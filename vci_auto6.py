import tkinter as tk
from tkinter import ttk, messagebox
from xml.etree.ElementTree import Element, SubElement, ElementTree
from datetime import datetime, timedelta
import re

# pyinstaller -w -F --add-binary="C:/Users/kod03/AppData/Local/Programs/Python/Python311/tcl/tkdnd2.8;tkdnd2.8" vci_auto6.py

def parse_excel_data(data):
    """
    엑셀 데이터를 파싱하여 컨테이너 정보로 변환
    """
    try:
        lines = data.strip().split('\n')
        container_data = []
        
        # 각 행 처리
        for row_idx, line in enumerate(lines):
            cells = line.strip().split('\t')
            
            # FULL 데이터 (첫 4열)
            for col_idx in range(1, min(4, len(cells))):
                try:
                    cell_value = cells[col_idx].strip()
                    count = int(cell_value) if cell_value else 0
                    if count > 0:
                        if row_idx == 0:  # LOCAL
                            container_type = "DIMP" if col_idx < 3 else "DEXP"
                        else:  # T/S
                            container_type = "TRAN"
                        
                        size = 20 if col_idx in [1, 3] else 40
                        container_data.append({
                            'type': container_type,
                            'operator': 'MSC',
                            'containersize': size,
                            'fullempty': 'F',
                            'number': count,
                            'terminal': 'PTSIEPS'
                        })
                except (ValueError, IndexError):
                    continue
                    
            # EMPTY 데이터 (후반 4열)
            for col_idx in range(4, min(8, len(cells))):
                try:
                    cell_value = cells[col_idx].strip()
                    count = int(cell_value) if cell_value else 0
                    if count > 0:
                        if row_idx == 0:  # LOCAL
                            container_type = "DIMP" if col_idx < 6 else "DEXP"
                        else:  # T/S
                            container_type = "TRAN"
                        
                        size = 20 if col_idx in [4, 6] else 40
                        container_data.append({
                            'type': container_type,
                            'operator': 'MSC',
                            'containersize': size,
                            'fullempty': 'E',
                            'number': count,
                            'terminal': 'PTSIEPS'
                        })
                except (ValueError, IndexError):
                    continue
        
        return container_data
    except Exception as e:
        print(f"Error in parse_excel_data: {e}")
        return []

class VCIXMLGenerator:
    def __init__(self):
        self.root = Element('vci')
        self.root.set('version', '1.0')
        self.root.set('type', 'VESSEL')

    def add_header(self, vessel, voyage, portun, master, berth):
        header = SubElement(self.root, 'header')
        vessel_elem = SubElement(header, 'vessel')
        vessel_elem.text = vessel
        voyage_elem = SubElement(header, 'voyage')
        voyage_elem.text = voyage
        portun_elem = SubElement(header, 'portun')
        portun_elem.text = portun
        master_elem = SubElement(header, 'master')
        master_elem.text = master
        berth_elem = SubElement(header, 'berth')
        berth_elem.text = berth

class VCIGeneratorGUI:
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.master.title("VCI XML Generator")
        
        # Initialize data storage
        self.gang_operation_data = []
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both')
        
        # Create tabs
        self.create_arrival_tab()
        self.create_operations_tab()
        self.create_departure_tab()
        self.create_discharge_tab()
        self.create_load_tab()
        self.create_shifting_tab()

    def handle_qc_paste(self, event=None):
        try:
            # Get clipboard data
            clipboard_data = self.clipboard_get()
            if not clipboard_data.strip():
                messagebox.showerror("Error", "Clipboard is empty")
                return

            # Parse clipboard data
            lines = clipboard_data.strip().split('\n')
            qc_data = []
            
            for line in lines:
                parts = line.strip().split('\t')
                if len(parts) >= 2:  # Ensure we have at least start and end time
                    try:
                        # Parse start and end times
                        start_time = datetime.strptime(parts[0].strip(), "%Y%m%d %H%M")
                        end_time = datetime.strptime(parts[1].strip(), "%Y%m%d %H%M")
                        qc_data.append((start_time, end_time))
                    except ValueError as e:
                        print(f"Error parsing time: {e}")
                        continue

            if not qc_data:
                messagebox.showerror("Error", "No valid QC data found in clipboard")
                return

            # Process QC data to find concurrent operations
            all_times = []
            for start, end in qc_data:
                all_times.append((start, 'start'))
                all_times.append((end, 'end'))
            
            all_times.sort(key=lambda x: (x[0], x[1] != 'start'))  # Sort by time, prioritizing 'start' events
            
            active_qcs = 0
            self.gang_operation_data = []  # Clear existing data
            
            for i in range(len(all_times) - 1):
                current_time, event = all_times[i]
                
                if event == 'start':
                    active_qcs += 1
                else:
                    active_qcs -= 1
                
                # Only create entries when there are active QCs
                if active_qcs > 0:
                    next_time = all_times[i + 1][0]
                    
                    # Add one minute to start time
                    adjusted_start = current_time + timedelta(minutes=1)
                    
                    # Create gang operation entry
                    gang_entry = {
                        'number': str(active_qcs),
                        'start_time': adjusted_start.strftime("%Y%m%d %H%M"),
                        'end_time': next_time.strftime("%Y%m%d %H%M")
                    }
                    self.gang_operation_data.append(gang_entry)
            
            messagebox.showinfo("Success", f"Processed {len(self.gang_operation_data)} gang operation entries")
            
        except Exception as e:
            error_msg = f"Error processing clipboard data: {str(e)}"
            messagebox.showerror("Error", error_msg)
            print(error_msg)  # For debugging

    def create_arrival_tab(self):
        # Timeline 섹션
        timeline_frame = ttk.LabelFrame(self.notebook, text="Timeline")
        timeline_frame.pack(pady=5, fill='x', padx=10)
        
        # BOWTHARR 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="Arrival Bowthruster", width=35).pack(side='left')
        self.bowtharr_var = tk.StringVar(value="arrival in order")
        frame2 = ttk.Frame(frame)
        frame2.pack(side='left')
        ttk.Radiobutton(frame2, text="arrival in order", variable=self.bowtharr_var, 
                       value="arrival in order").pack(side='left', padx=5)
        ttk.Radiobutton(frame2, text="arrival out of order", variable=self.bowtharr_var,
                       value="arrival out of order").pack(side='left', padx=5)

        # 날짜/시간 입력 필드들
        time_fields = [
         #   ("PILNOT", "pilot_time_entry"),
          #  ("PILORDFOR", "pilord_time_entry"),
            ("Vessel Arrives At Pilot Station/Roads", "arrpil_time_entry"),
           # ("FIRLINASH", "firline_time_entry")
        ]
        
        for label_text, entry_name in time_fields:
            self.create_datetime_entry(timeline_frame, label_text, entry_name)

        # Arrival 탭의 REAFORANC 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="Lost Time Waiting For A Berth", width=30).pack(side='left')
        self.arr_reaforanc_var = tk.StringVar(value="Less than 1 hour")  # arr_ 접두어 추가
        ttk.Combobox(frame, textvariable=self.arr_reaforanc_var,
                    values=["Less than 1 hour","Berth congestion - On window","Berth Congestion - Off window","Bunkering","Harbour traffic","Authorities","Quarantine","Arrived ahead of schedule","To avoid additional pilotage costs","To avoid additional towage costs","To avoid additional terminal costs","To avoid other port costs","Vessel repairs","Preferred berthing","Waiting for Transhipment Cargo","Navigation Restriction","Public Holidays","Balast/debalast","Geneva instructions","Other"], 
                    state="readonly", width=30).pack(side='left', padx=40)

        # Draft 섹션
        draft_frame = ttk.LabelFrame(self.notebook, text="Draft")
        draft_frame.pack(pady=5, fill='x', padx=10)
        
        draft_fields = [("AFT", "arr_draft_aft_entry"), ("FWD", "arr_draft_fwd_entry")]
        for label_text, entry_name in draft_fields:
            frame = ttk.Frame(draft_frame)
            frame.pack(pady=5, fill='x')
            ttk.Label(frame, text=label_text, width=30).pack(side='left')
            entry = ttk.Entry(frame, width=6)  # 흘수값 (예: 11.45)
            entry.pack(side='left')
            setattr(self, entry_name, entry)

        # Pilots 섹션
        pilots_frame = ttk.LabelFrame(self.notebook, text="Pilots")
        pilots_frame.pack(pady=5, fill='x', padx=10)
        
        pilot_fields = [("From", "arr_pilot_from_entry"), ("To", "arr_pilot_to_entry")]
        for label_text, entry_name in pilot_fields:
            self.create_datetime_entry(pilots_frame, label_text, entry_name)

        # Towages 섹션
        towages_frame = ttk.LabelFrame(self.notebook, text="Towages")
        towages_frame.pack(pady=5, fill='x', padx=70)
        
        # Add Tug 버튼 추가
        add_tug_button = ttk.Button(towages_frame, text="Add Tug",
                                   command=lambda: self.add_tug_frame(towages_frame, 'arrival'))
        add_tug_button.pack(pady=5)
        
        self.arr_tug_entries = []
        # 초기 1개의 고정 Tug 생성
        tug_frame = ttk.LabelFrame(towages_frame, text="Tug 1")
        tug_frame.pack(pady=5, fill='x')
        
        entries = {}
        # From/To 필드는 datetime entry로 생성
        self.create_datetime_entry(tug_frame, "From", "arr_tug1_from_entry")
        self.create_datetime_entry(tug_frame, "To", "arr_tug1_to_entry")
        
        # Comment 필드는 일반 entry로 생성
        comment_frame = ttk.Frame(tug_frame)
        comment_frame.pack(pady=5, fill='x')
        ttk.Label(comment_frame, text="Comment", width=30).pack(side='left')
        comment_entry = ttk.Entry(comment_frame, width=30)
        comment_entry.pack(side='left', padx=40)
        setattr(self, "arr_tug1_comment_entry", comment_entry)
        
        entries["From"] = getattr(self, "arr_tug1_from_entry")
        entries["To"] = getattr(self, "arr_tug1_to_entry")
        entries["Comment"] = comment_entry
        entries["Frame"] = tug_frame
        self.arr_tug_entries.append(entries)

    def add_tug_frame(self, parent_frame, tab_type):
        """새로운 Tug 프레임 추가 함수"""
        tug_count = len(self.arr_tug_entries if tab_type == 'arrival' else self.dep_tug_entries) + 1
        
        tug_frame = ttk.LabelFrame(parent_frame, text=f"Tug {tug_count}")
        tug_frame.pack(pady=5, fill='x')
        
        entries = {}
        # From/To 필드는 datetime entry로 생성
        from_entry = self.create_datetime_entry(tug_frame, "From")
        to_entry = self.create_datetime_entry(tug_frame, "To")
        
        # Comment 필드는 일반 entry로 생성
        comment_frame = ttk.Frame(tug_frame)
        comment_frame.pack(pady=5, fill='x')
        ttk.Label(comment_frame, text="Comment", width=30).pack(side='left')
        comment_entry = ttk.Entry(comment_frame, width=30)
        comment_entry.pack(side='left', padx=40)
        
        # Delete 버튼 추가 (추가된 Tug에만)
        delete_button = ttk.Button(tug_frame, text="Delete",
                                 command=lambda: self.delete_tug_frame(tug_frame, tab_type))
        delete_button.pack(pady=5)
        
        entries["From"] = from_entry
        entries["To"] = to_entry
        entries["Comment"] = comment_entry
        entries["Frame"] = tug_frame
        
        if tab_type == 'arrival':
            self.arr_tug_entries.append(entries)
        else:
            self.dep_tug_entries.append(entries)

    def delete_tug_frame(self, frame, tab_type):
        """Tug 프레임 삭제 함수"""
        frame.destroy()
        if tab_type == 'arrival':
            self.arr_tug_entries = [entry for entry in self.arr_tug_entries 
                                   if entry["Frame"] != frame]
        else:
            self.dep_tug_entries = [entry for entry in self.dep_tug_entries 
                                   if entry["Frame"] != frame]

    def create_operations_tab(self):
        # Gang Summary 섹션 (하단)
        gang_summary_frame = ttk.LabelFrame(self.notebook, text="Gang Summary")
        gang_summary_frame.pack(pady=5, fill='x', padx=70)

        # Timeline 섹션 (상단)
        timeline_frame = ttk.LabelFrame(self.notebook, text="Timeline")
        timeline_frame.pack(pady=5, fill='x', padx=70)

        # Gang Operation 섹션 추가
        gang_operation_frame = ttk.LabelFrame(self.notebook, text="Gang Operation")
        gang_operation_frame.pack(pady=5, fill='x', padx=70)

        # Gang Operation 라인을 저장할 프레임
        self.gang_operation_lines_frame = ttk.Frame(gang_operation_frame)
        self.gang_operation_lines_frame.pack(pady=5, fill='x')
        
        # Gang Operation 라인을 저장할 리스트
        self.gang_operation_lines = []

        # Add Line 버튼 추가
        add_button = ttk.Button(gang_operation_frame, text="Add Line",
                              command=lambda: self.add_gang_operation_line())
        add_button.pack(pady=5)

        # 붙여넣기 이벤트 바인딩
        gang_operation_frame.bind('<Control-v>', lambda e: self.handle_qc_paste())
        self.gang_operation_lines_frame.bind('<Control-v>', lambda e: self.handle_qc_paste())

        # Operations Start - 먼저 생성
        frame = ttk.Frame(gang_summary_frame)
        frame.pack(pady=5, fill='x')
        self.operationsstart_entry = self.create_datetime_entry(frame, "Operations Start")
        self.operationsstart_entry.pack(side='left', padx=40)

        # Gang Start - datetime entry 생성 시 자동으로 all_datetime_entries에 추가됨
        frame = ttk.Frame(gang_summary_frame)
        frame.pack(pady=5, fill='x')
        self.gangstart_entry = self.create_datetime_entry(frame, "Gang Start")
        self.gangstart_entry.pack(side='left', padx=40)
        
        # operationsstart_entry 값이 변경될 때마다 gangstart_entry 값을 10분 전으로 설정
        def update_gang_start(*args):
            try:
                # 입력된 값 가져오기
                ops_start_str = self.operationsstart_entry.get()
                if not ops_start_str:
                    return
                
                # 입력 형식에 따라 파싱 (YYYYMMDD HHMM 또는 YYYY-MM-DDThh:mm:ss)
                if len(ops_start_str) == 13 and ' ' in ops_start_str:  # YYYYMMDD HHMM 형식
                    date_part = ops_start_str[:8]
                    time_part = ops_start_str[9:]
                    year = date_part[:4]
                    month = date_part[4:6]
                    day = date_part[6:8]
                    hour = time_part[:2]
                    minute = time_part[2:4]
                    ops_start = datetime(int(year), int(month), int(day), int(hour), int(minute))
                else:  # ISO 형식
                    ops_start = datetime.strptime(ops_start_str, "%Y-%m-%dT%H:%M:%S")
                
                # 10분 전 시간 계산
                gang_start = ops_start - timedelta(minutes=10)
                
                # gangstart_entry에 설정
                self.gangstart_entry.delete(0, tk.END)
                self.gangstart_entry.insert(0, gang_start.strftime("%Y%m%d %H%M"))
                
                # Operations Commence(opecom_entry)에도 같은 값 설정
                self.opecom_entry.delete(0, tk.END)
                self.opecom_entry.insert(0, ops_start_str)
                
                print(f"Updated gang start to: {gang_start.strftime('%Y%m%d %H%M')}")
                print(f"Updated operations commence to: {ops_start_str}")
            except Exception as e:
                print(f"Error updating gang start: {str(e)}")
                
        # 이벤트 바인딩 - KeyRelease와 FocusOut 모두 바인딩
        self.operationsstart_entry.bind('<KeyRelease>', update_gang_start)
        self.operationsstart_entry.bind('<FocusOut>', update_gang_start)

        # Gang Finish
        frame = ttk.Frame(gang_summary_frame)
        frame.pack(pady=5, fill='x')
        self.gangfinish_entry = self.create_datetime_entry(frame, "Gang Finish")
        self.gangfinish_entry.pack(side='left', padx=40)

        # Operations Finish
        frame = ttk.Frame(gang_summary_frame)
        frame.pack(pady=5, fill='x')
        self.operationsfinish_entry = self.create_datetime_entry(frame, "Operations Finish")
        self.operationsfinish_entry.pack(side='left', padx=40)
        
        # operationsfinish_entry 값이 변경될 때마다 gangfinish_entry 값을 10분 후로 설정
        def update_gang_finish(*args):
            try:
                # 입력된 값 가져오기
                ops_finish_str = self.operationsfinish_entry.get()
                if not ops_finish_str:
                    return
                
                # 입력 형식에 따라 파싱 (YYYYMMDD HHMM 또는 YYYY-MM-DDThh:mm:ss)
                if len(ops_finish_str) == 13 and ' ' in ops_finish_str:  # YYYYMMDD HHMM 형식
                    date_part = ops_finish_str[:8]
                    time_part = ops_finish_str[9:]
                    year = date_part[:4]
                    month = date_part[4:6]
                    day = date_part[6:8]
                    hour = time_part[:2]
                    minute = time_part[2:4]
                    ops_finish = datetime(int(year), int(month), int(day), int(hour), int(minute))
                else:  # ISO 형식
                    ops_finish = datetime.strptime(ops_finish_str, "%Y-%m-%dT%H:%M:%S")
                
                # 10분 후 시간 계산
                gang_finish = ops_finish + timedelta(minutes=10)
                
                # gangfinish_entry에 설정
                self.gangfinish_entry.delete(0, tk.END)
                self.gangfinish_entry.insert(0, gang_finish.strftime("%Y%m%d %H%M"))
                
                # Operations Completed(opecomp_entry)에도 같은 값 설정
                self.opecomp_entry.delete(0, tk.END)
                self.opecomp_entry.insert(0, ops_finish_str)
                
                # Lashing Completed(lascomp_entry)에 10분 후 값 설정
                lashing_completed = ops_finish + timedelta(minutes=10)
                self.lascomp_entry.delete(0, tk.END)
                self.lascomp_entry.insert(0, lashing_completed.strftime("%Y%m%d %H%M"))
                
                print(f"Updated gang finish to: {gang_finish.strftime('%Y%m%d %H%M')}")
                print(f"Updated operations completed to: {ops_finish_str}")
                print(f"Updated lashing completed to: {lashing_completed.strftime('%Y%m%d %H%M')}")
            except Exception as e:
                print(f"Error updating gang finish: {str(e)}")
                
        # 이벤트 바인딩 - KeyRelease와 FocusOut 모두 바인딩
        self.operationsfinish_entry.bind('<KeyRelease>', update_gang_finish)
        self.operationsfinish_entry.bind('<FocusOut>', update_gang_finish)

        # DOCSIDTO 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="Docked Side To [Port / Starboard]", width=35).pack(side='left')
        self.docsidto_var = tk.StringVar(value="Port")
        
        radio_frame = ttk.Frame(frame)
        radio_frame.pack(side='left')
        ttk.Radiobutton(radio_frame, text="Port", variable=self.docsidto_var, 
                       value="Port").pack(side='left', padx=5)
        ttk.Radiobutton(radio_frame, text="Starboard", variable=self.docsidto_var,
                       value="Starboard").pack(side='left', padx=5)

        # 날짜/시간 입력 필드들
        time_fields = [
            ("Docked [All Fast] At Terminal", "docatter_entry"),
            ("Gangway Down", "ganwaydown_entry"),
            ("Operations Commence", "opecom_entry"),
            ("Operations Completed", "opecomp_entry"),
            ("Lashing Completed", "lascomp_entry")
        ]
        
        for label_text, entry_name in time_fields:
            self.create_datetime_entry(timeline_frame, label_text, entry_name)
            
        # Docked [All Fast] At Terminal 값이 변경될 때마다 Gangway Down 값을 30분 후로 설정
        def update_gangway_down(*args):
            try:
                # 입력된 값 가져오기
                docatter_str = self.docatter_entry.get()
                if not docatter_str:
                    return
                
                # 입력 형식에 따라 파싱 (YYYYMMDD HHMM 또는 YYYY-MM-DDThh:mm:ss)
                if len(docatter_str) == 13 and ' ' in docatter_str:  # YYYYMMDD HHMM 형식
                    date_part = docatter_str[:8]
                    time_part = docatter_str[9:]
                    year = date_part[:4]
                    month = date_part[4:6]
                    day = date_part[6:8]
                    hour = time_part[:2]
                    minute = time_part[2:4]
                    docatter = datetime(int(year), int(month), int(day), int(hour), int(minute))
                else:  # ISO 형식
                    docatter = datetime.strptime(docatter_str, "%Y-%m-%dT%H:%M:%S")
                
                # 30분 후 시간 계산
                gangway_down = docatter + timedelta(minutes=30)
                
                # ganwaydown_entry에 설정
                self.ganwaydown_entry.delete(0, tk.END)
                self.ganwaydown_entry.insert(0, gangway_down.strftime("%Y%m%d %H%M"))
                
                print(f"Updated gangway down to: {gangway_down.strftime('%Y%m%d %H%M')}")
            except Exception as e:
                print(f"Error updating gangway down: {str(e)}")
                
        # 이벤트 바인딩 - KeyRelease와 FocusOut 모두 바인딩
        self.docatter_entry.bind('<KeyRelease>', update_gangway_down)
        self.docatter_entry.bind('<FocusOut>', update_gangway_down)

        # LASCOMPBY 라디오버튼
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="Lashing Done By", width=35).pack(side='left')
        self.lascompby_var = tk.StringVar(value="Terminal")
        
        radio_frame = ttk.Frame(frame)
        radio_frame.pack(side='left')
        ttk.Radiobutton(radio_frame, text="Terminal", variable=self.lascompby_var, 
                       value="Terminal").pack(side='left', padx=5)
        ttk.Radiobutton(radio_frame, text="Vessel Crew", variable=self.lascompby_var,
                       value="Vessel Crew").pack(side='left', padx=5)

    def add_gang_operation_line(self, number="", start_time="", end_time=""):
        """Gang Operation 라인 추가"""
        frame = ttk.Frame(self.gang_operation_lines_frame)
        frame.pack(pady=5, fill='x')
        
        # Number 입력
        number_entry = ttk.Entry(frame, width=8)
        if number:
            number_entry.insert(0, str(number))
        number_entry.pack(side='left', padx=2)
        
        # Start Time 입력
        start_time_entry = ttk.Entry(frame, width=15)
        if start_time:
            start_time_entry.insert(0, start_time)
        start_time_entry.pack(side='left', padx=2)
        
        # End Time 입력
        end_time_entry = ttk.Entry(frame, width=15)
        if end_time:
            end_time_entry.insert(0, end_time)
        end_time_entry.pack(side='left', padx=2)
        
        # 삭제 버튼
        delete_button = ttk.Button(frame, text="X",
                                 command=lambda: self.delete_gang_operation_line(frame))
        delete_button.pack(side='left', padx=2)
        
        line_data = {
            "frame": frame,
            "number": number_entry,
            "start_time": start_time_entry,
            "end_time": end_time_entry
        }
        
        self.gang_operation_lines.append(line_data)
        return line_data

    def delete_gang_operation_line(self, frame):
        """Gang Operation 라인 삭제"""
        frame.destroy()
        self.gang_operation_lines = [line for line in self.gang_operation_lines 
                                   if line["frame"] != frame]

    def create_departure_tab(self):
        # Timeline 섹션
        timeline_frame = ttk.LabelFrame(self.notebook, text="Timeline")
        timeline_frame.pack(pady=5, fill='x', padx=70)
        
        # BOWTHDEP 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="Departure Bowthruster", width=30).pack(side='left')
        self.bowthdep_var = tk.StringVar(value="departure in order")
        frame2 = ttk.Frame(frame)
        frame2.pack(side='left')
        ttk.Radiobutton(frame2, text="departure in order", variable=self.bowthdep_var, 
                       value="departure in order").pack(side='left', padx=5)
        ttk.Radiobutton(frame2, text="departure out of order", variable=self.bowthdep_var,
                       value="departure out of order").pack(side='left', padx=5)

        # Lost Time Waiting To Sail 레이블과 REAFORANC 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="Lost Time Waiting To Sail", width=30).pack(side='left')
        self.dep_reaforanc_var = tk.StringVar(value="Less than 1 hour")
        ttk.Combobox(frame, textvariable=self.dep_reaforanc_var,
                    values=["Less than 1 hour","Berth congestion - On window","Berth Congestion - Off window","Bunkering","Harbour traffic","Authorities","Quarantine","Arrived ahead of schedule","To avoid additional pilotage costs","To avoid additional towage costs","To avoid additional terminal costs","To avoid other port costs","Vessel repairs","Preferred berthing","Waiting for Transhipment Cargo","Navigation Restriction","Public Holidays","Balast/debalast","Geneva instructions","Other"], 
                    state="readonly", width=30).pack(side='left', padx=40)

        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        self.vesundoc_entry = self.create_datetime_entry(frame, "Sailed From Berth")
        self.vesundoc_entry.pack(side='left', padx=40)

        # Draft 섹션
        draft_frame = ttk.LabelFrame(self.notebook, text="Draft")
        draft_frame.pack(pady=5, fill='x', padx=10)
        
        draft_fields = [("AFT", "dep_draft_aft_entry"), ("FWD", "dep_draft_fwd_entry")]
        for label_text, entry_name in draft_fields:
            frame = ttk.Frame(draft_frame)
            frame.pack(pady=5, fill='x')
            ttk.Label(frame, text=label_text, width=15).pack(side='left')
            entry = ttk.Entry(frame, width=6)  # 흘수값 (예: 11.45)
            entry.pack(side='left')
            setattr(self, entry_name, entry)

        # Pilots 섹션
        pilots_frame = ttk.LabelFrame(self.notebook, text="Pilots")
        pilots_frame.pack(pady=5, fill='x', padx=10)
        
        pilot_fields = [("From", "dep_pilot_from_entry"), ("To", "dep_pilot_to_entry")]
        for label_text, entry_name in pilot_fields:
            self.create_datetime_entry(pilots_frame, label_text, entry_name)

        # Towages 섹션
        towages_frame = ttk.LabelFrame(self.notebook, text="Towages")
        towages_frame.pack(pady=5, fill='x', padx=70)
        
        # Add Tug 버튼 추가
        add_tug_button = ttk.Button(towages_frame, text="Add Tug",
                                   command=lambda: self.add_tug_frame(towages_frame, 'departure'))
        add_tug_button.pack(pady=5)
        
        self.dep_tug_entries = []
        # 초기 1개의 고정 Tug 생성
        tug_frame = ttk.LabelFrame(towages_frame, text="Tug 1")
        tug_frame.pack(pady=5, fill='x')
        
        entries = {}
        # From/To 필드는 datetime entry로 생성
        self.create_datetime_entry(tug_frame, "From", "dep_tug1_from_entry")
        self.create_datetime_entry(tug_frame, "To", "dep_tug1_to_entry")
        
        # Comment 필드는 일반 entry로 생성
        comment_frame = ttk.Frame(tug_frame)
        comment_frame.pack(pady=5, fill='x')
        ttk.Label(comment_frame, text="Comment", width=30).pack(side='left')
        comment_entry = ttk.Entry(comment_frame, width=30)
        comment_entry.pack(side='left', padx=40)
        setattr(self, "dep_tug1_comment_entry", comment_entry)
        
        entries["From"] = getattr(self, "dep_tug1_from_entry")
        entries["To"] = getattr(self, "dep_tug1_to_entry")
        entries["Comment"] = comment_entry
        entries["Frame"] = tug_frame
        self.dep_tug_entries.append(entries)

    def create_discharge_tab(self):
        # 컨테이너 라인 프레임
        self.discharge_lines_frame = ttk.Frame(self.notebook)
        self.discharge_lines_frame.pack(pady=5, fill='x', padx=10)
        
        # 초기 컨테이너 라인은 생성하지 않음
        self.discharge_lines = []

    def create_load_tab(self):
        # 컨테이너 라인 프레임
        self.load_lines_frame = ttk.Frame(self.notebook)
        self.load_lines_frame.pack(pady=5, fill='x', padx=10)
        
        # 초기 컨테이너 라인은 생성하지 않음
        self.load_lines = []

    def add_container_line_with_data(self, tab_type, container_data):
        """
        컨테이너 데이터로 라인 추가
        """
        frame = ttk.Frame(self.discharge_lines_frame if tab_type == "discharge" 
                         else self.load_lines_frame)
        frame.pack(pady=5, fill='x')
        
        # 타입 선택
        type_values = ["DIMP", "DEXP", "TRAN"]
        type_var = tk.StringVar(value=container_data['type'])
        type_combo = ttk.Combobox(frame, textvariable=type_var,
                                 values=type_values, state="readonly", width=8)
        type_combo.pack(side='left', padx=2)
        
        # 기타 필드들
        operator_var = tk.StringVar(value=container_data['operator'])
        ttk.Combobox(frame, textvariable=operator_var,
                    values=["MSC","HLC","ONE","YML","MSK","HPL","MAE","HMM"], 
                    state="readonly", width=5).pack(side='left', padx=2)
        
        size_var = tk.StringVar(value=str(container_data['containersize']))
        ttk.Combobox(frame, textvariable=size_var,
                    values=["20", "40"], state="readonly", width=4).pack(side='left', padx=2)
        
        fe_var = tk.StringVar(value=container_data['fullempty'])
        ttk.Combobox(frame, textvariable=fe_var,
                    values=["F", "E"], state="readonly", width=3).pack(side='left', padx=2)
        
        number_entry = ttk.Entry(frame, width=6)
        number_entry.insert(0, str(container_data['number']))
        number_entry.pack(side='left', padx=2)
        
        # 삭제 버튼
        delete_button = ttk.Button(frame, text="X",
                                 command=lambda: self.delete_container_line(frame, tab_type))
        delete_button.pack(side='left', padx=2)
        
        line_data = {
            "frame": frame,
            "type": type_var,
            "operator": operator_var,
            "size": size_var,
            "fe": fe_var,
            "number": number_entry
        }
        
        if tab_type == "discharge":
            self.discharge_lines.append(line_data)
        else:
            self.load_lines.append(line_data)

    def delete_container_line(self, frame, tab_type):
        frame.destroy()
        if tab_type == "discharge":
            self.discharge_lines = [line for line in self.discharge_lines 
                                  if line["frame"] != frame]
        else:
            self.load_lines = [line for line in self.load_lines 
                             if line["frame"] != frame]

    def create_shifting_tab(self):
        # Lid Moves 섹션
        lid_frame = ttk.LabelFrame(self.notebook, text="Lid Moves")
        lid_frame.pack(pady=5, fill='x', padx=10)
        
        lid_fields = [("On", "lid_on_entry"), ("Off", "lid_off_entry")]
        for label_text, entry_name in lid_fields:
            frame = ttk.Frame(lid_frame)
            frame.pack(pady=5, fill='x')
            ttk.Label(frame, text=label_text, width=15).pack(side='left')
            entry = ttk.Entry(frame, width=3)  # 2자리 정수값
            entry.pack(side='left')
            setattr(self, entry_name, entry)

        add_button = ttk.Button(self.notebook, text="Add Shifting Line",
                              command=self.add_shifting_line)
        add_button.pack(pady=5)
    
        # Container Shifting 섹션
        self.shifting_lines_frame = ttk.LabelFrame(self.notebook, text="Container Shifting")
        self.shifting_lines_frame.pack(pady=5, fill='x', padx=10)

        # Header labels - 한 번만 표시
        header_frame = ttk.Frame(self.shifting_lines_frame)
        header_frame.pack(fill='x')
        
        headers = ["Account", "Type", "Size", "#", "F/E", "OOG", "RF", "IMO", "Reason"]
        widths = [11, 10, 7, 5, 7, 6, 5, 5, 50]
        
        for header, width in zip(headers, widths):
            ttk.Label(header_frame, text=header, width=width).pack(side='left', padx=2)

        self.shifting_lines = []
        self.add_shifting_line()  # 초기 라인 하나 추가

    def add_shifting_line(self):
        frame = ttk.Frame(self.shifting_lines_frame)
        frame.pack(pady=5, fill='x')
        
        # Account
        account_var = tk.StringVar(value="MSCU")
        ttk.Combobox(frame, textvariable=account_var,
                    values=["MSCU","ZIMU","HDMU","HLCU","MAEU"], state="readonly", width=8).pack(side='left', padx=2)
        
        # Type
        type_var = tk.StringVar(value="Restow")
        ttk.Combobox(frame, textvariable=type_var,
                    values=["Restow"], state="readonly", width=8).pack(side='left', padx=2)
        
        # Container Size
        size_var = tk.StringVar(value="40")
        ttk.Combobox(frame, textvariable=size_var,
                    values=["20", "40"], state="readonly", width=4).pack(side='left', padx=2)
        
        # Value
        value_entry = ttk.Entry(frame, width=4)  # 컨테이너 수량
        value_entry.pack(side='left', padx=2)
        
        # Full/Empty
        fe_var = tk.StringVar(value="F")
        ttk.Combobox(frame, textvariable=fe_var,
                    values=["F", "E"], state="readonly", width=3).pack(side='left', padx=10)
        
        # OOG
        oog_var = tk.StringVar(value="0")
        oog_check = ttk.Checkbutton(frame, variable=oog_var, onvalue="1", offvalue="0")
        oog_check.pack(side='left', padx=10)
        
        # Reefer
        reefer_var = tk.StringVar(value="0")
        reefer_check = ttk.Checkbutton(frame, variable=reefer_var, onvalue="1", offvalue="0")
        reefer_check.pack(side='left', padx=10)
        
        # IMO
        imo_var = tk.StringVar(value="0") 
        imo_check = ttk.Checkbutton(frame, variable=imo_var, onvalue="1", offvalue="0")
        imo_check.pack(side='left', padx=10)
        
        # Reason
        reason_var = tk.StringVar(value="Restow of optional cargo onboard, to maximize vsl capacity")
        ttk.Combobox(frame, textvariable=reason_var,
                    values=["Restow of optional cargo onboard, to maximize vsl capacity"], 
                    state="readonly", width=50).pack(side='left', padx=10)
        
        # Not For MSC Account
        notformscaccount_var = tk.BooleanVar()
        notformscaccount_check = ttk.Checkbutton(frame, text="Not For MSC", variable=notformscaccount_var)
        notformscaccount_check.pack(side='left', padx=10)
        
        # Delete 버튼
        delete_button = ttk.Button(frame, text="Delete",
                                 command=lambda: self.delete_shifting_line(frame))
        delete_button.pack(side='left', padx=2)
        
        line_data = {
            "frame": frame,
            "account": account_var,
            "type": type_var,
            "size": size_var,
            "value": value_entry,
            "fe": fe_var,
            "oog": oog_var,
            "reefer": reefer_var,
            "imo": imo_var,
            "reason": reason_var,
            "notformscaccount": notformscaccount_var
        }
        
        self.shifting_lines.append(line_data)

    def delete_shifting_line(self, frame):
        frame.destroy()
        self.shifting_lines = [line for line in self.shifting_lines 
                              if line["frame"] != frame]

    def convert_datetime(self, input_str):
        try:
            date_part = input_str[:8]
            time_part = input_str[9:]
            formatted_date = f"{date_part[:4]}-{date_part[4:6]}-{date_part[6:]}"
            formatted_time = f"{time_part[:2]}:{time_part[2:]}:00"
            return f"{formatted_date}T{formatted_time}"
        except:
            return input_str

    def convert_to_iso_datetime(self, datetime_str):
        if not datetime_str:
            return ""
        try:
            # YYYYMMDD HHMM 형식을 YYYY-MM-DDThh:mm:ss 형식으로 변환
            date_part = datetime_str[:8]
            time_part = datetime_str[-4:] if len(datetime_str) > 8 else "0000"
            
            year = date_part[:4]
            month = date_part[4:6]
            day = date_part[6:8]
            hour = time_part[:2]
            minute = time_part[2:4]
            
            return f"{year}-{month}-{day}T{hour}:{minute}:00"
        except:
            return ""

    def generate_xml(self):
        try:
            # Create XML generator
            generator = VCIXMLGenerator()
            
            # Add header information
            generator.add_header(
                vessel=self.vessel_entry.get(),
                voyage=self.voyage_entry.get(),
                portun=self.portun_entry.get(),
                master=self.master_entry.get(),
                berth=self.berth_entry.get()
            )
            
            # Add operations section
            operations = SubElement(generator.root, 'operations')
            
            # Add gang work time section if gang operation data exists
            if hasattr(self, 'gang_operation_data') and self.gang_operation_data:
                gangworktime = SubElement(operations, 'gangworktime')
                
                for gang_data in self.gang_operation_data:
                    gang = SubElement(gangworktime, 'gang')
                    
                    # Set required attributes according to schema
                    gang.set('number', gang_data['number'])
                    
                    # Convert YYYYMMDD HHMM to YYYY-MM-DDThh:mm:ss format
                    start_dt = datetime.strptime(gang_data['start_time'], "%Y%m%d %H%M")
                    end_dt = datetime.strptime(gang_data['end_time'], "%Y%m%d %H%M")
                    
                    gang.set('from', start_dt.strftime("%Y-%m-%dT%H:%M:%S"))
                    gang.set('to', end_dt.strftime("%Y-%m-%dT%H:%M:%S"))
                    gang.set('workingperiodpayrate', 'NORMAL')  # Required by schema
                    gang.set('terminal', 'PNIT')  # Required by schema
            
            # Create and save XML file
            tree = ElementTree(generator.root)
            tree.write('output.xml', encoding='utf-8', xml_declaration=True)
            messagebox.showinfo("Success", "XML file has been generated successfully!")
            
        except Exception as e:
            error_msg = f"Error generating XML: {str(e)}"
            messagebox.showerror("Error", error_msg)
            print(f"Debug - Error details: {str(e)}")  # For debugging

    def add_timeline_field(self, parent, remarkscode, type_, value):
        field = SubElement(parent, 'field')
        field.set('remarkscode', remarkscode)
        field.set('type', type_)
        field.set('value', value)

    def add_draft_item(self, parent, type_, value):
        item = SubElement(parent, 'item')
        item.set('type', type_)
        item.set('value', str(value))
        item.set('unit', 'm')

    def handle_paste(self, tab_type):
        """
        클립보드 데이터 처리
        """
        try:
            clipboard_data = self.root.clipboard_get()
            if not clipboard_data.strip():
                messagebox.showwarning("Warning", "클립보드가 비어있습니다.")
                return
                
            container_data = parse_excel_data(clipboard_data)
            if not container_data:
                messagebox.showwarning("Warning", "파싱할 수 있는 컨테이너 데이터가 없습니다.")
                return
            
            # 기존 컨테이너 라인 삭제
            if tab_type == 'discharge':
                for line in self.discharge_lines:
                    line["frame"].destroy()
                self.discharge_lines.clear()
            else:
                for line in self.load_lines:
                    line["frame"].destroy()
                self.load_lines.clear()
            
            # 새 데이터로 컨테이너 라인 생성
            for container in container_data:
                if tab_type == 'discharge':
                    self.add_container_line_with_data('discharge', container)
                else:
                    self.add_container_line_with_data('load', container)
                    
            messagebox.showinfo("Success", f"{len(container_data)}개의 컨테이너 데이터가 추가되었습니다.")
                    
        except Exception as e:
            error_msg = f"클립보드 데이터 처리 중 오류 발생: {str(e)}"
            messagebox.showerror("Error", error_msg)

def main():
    root = tk.Tk()
    app = VCIGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
