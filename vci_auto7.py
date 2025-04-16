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
        # 최상위 요소 생성
        self.root = Element('vcidata')
        self.root.set('version', '5')
        self.root.set('revision', '0')

    def add_header(self, vessel, voyage, portun, master, berth):
        header = SubElement(self.root, 'header')
        SubElement(header, 'vessel').text = vessel
        SubElement(header, 'voyage').text = voyage
        SubElement(header, 'portun').text = portun
        SubElement(header, 'master').text = master
        SubElement(header, 'berth').text = str(berth)

class VCIGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("VCI XML Generator")
        
        # 현재 날짜값을 저장할 변수 추가
        self.current_date = ""
        
        # 모든 날짜/시간 입력 필드들을 순서대로 저장할 리스트
        self.all_datetime_entries = []
        
        # 윈도우 크기 설정
        window_width = 400   # 원하는 너비
        window_height = 900   # 원하는 높이
        
        # 화면 중앙에 위치시키기
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        
        # 윈도우 크기와 위치 설정
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # 최소 윈도우 크기 설정
        self.root.minsize(1200, 700)
        
        # 생성 버튼을 상단에 배치
        self.generate_button = ttk.Button(root, text="Generate XML", command=self.generate_xml)
        self.generate_button.pack(pady=10)
        
        # 노트북(탭) 생성
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, expand=True, fill='both')  # fill='both'를 추가하여 가로세로 모두 채우기
        
        # 각 섹션별 탭 생성
        self.header_tab = ttk.Frame(self.notebook)
        self.arrival_tab = ttk.Frame(self.notebook)
        self.operations_tab = ttk.Frame(self.notebook)
        self.departure_tab = ttk.Frame(self.notebook)
        self.discharge_tab = ttk.Frame(self.notebook)
        self.load_tab = ttk.Frame(self.notebook)
        self.shifting_tab = ttk.Frame(self.notebook)
        
        # 탭 추가
        self.notebook.add(self.header_tab, text="Header")
        self.notebook.add(self.arrival_tab, text="Arrival")
        self.notebook.add(self.operations_tab, text="Operations")
        self.notebook.add(self.departure_tab, text="Departure")
        self.notebook.add(self.discharge_tab, text="Discharge")
        self.notebook.add(self.load_tab, text="Load")
        self.notebook.add(self.shifting_tab, text="Shifting")
        
        # 각 탭의 내용 초기화
        self.setup_header_tab()
        self.setup_arrival_tab()
        self.setup_operations_tab()
        self.setup_departure_tab()
        self.setup_discharge_tab()
        self.setup_load_tab()
        self.setup_shifting_tab()

        # 클립보드 바인딩 추가
        self.discharge_tab.bind('<Control-v>', lambda e: self.handle_paste('discharge'))
        self.load_tab.bind('<Control-v>', lambda e: self.handle_paste('load'))

        # 컨테이너 라인 엔트리 초기화
        self.discharge_entries = []
        self.load_entries = []

    def setup_header_tab(self):
        fields = [
            ("Vessel", "vessel_entry", 20),    # 선박명
            ("Voyage", "voyage_entry", 20),     # 항해번호
            ("Port UN", "portun_entry", 20)      # 항구코드
        ]
        
        for i, (label_text, entry_name, width) in enumerate(fields):
            frame = ttk.Frame(self.header_tab)
            frame.pack(pady=5, fill='x', padx=10)
            
            label = ttk.Label(frame, text=label_text, width=15)
            label.pack(side='left')
            
            entry = ttk.Entry(frame, width=width)  # 너비 지정
            entry.pack(side='left')
            setattr(self, entry_name, entry)

    def create_datetime_entry(self, parent, label_text, entry_name=None):
        """날짜/시간 입력 필드 생성 함수"""
        frame = ttk.Frame(parent)
        frame.pack(pady=5, fill='x')
        
        # 레이블
        ttk.Label(frame, text=label_text, width=30).pack(side='left')
        
        # 입력 필드
        entry = ttk.Entry(frame, width=13)
        entry.pack(side='left', padx=40)
        
        # 입력 필드에 이벤트 바인딩
        entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(entry))
        entry.bind('<KeyRelease>', lambda e: self.handle_date_change(entry))
        
        # 전체 날짜/시간 입력 필드 리스트에 순서대로 추가
        self.all_datetime_entries.append(entry)
        
        ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)
        
        if entry_name:
            setattr(self, entry_name, entry)
        
        return entry

    def handle_entry_focus(self, entry):
        """입력 필드 포커스 처리"""
        if not entry.get():
            # 비어있는 경우 현재 날짜 자동 입력
            if self.current_date:
                entry.insert(0, self.current_date + " ")
                entry.icursor(9)  # 커서를 시간 입력 위치로
        else:
            # 이미 값이 있는 경우 시간 부분으로 커서 이동
            entry.icursor(9)

    def handle_date_change(self, entry):
        """날짜 변경 처리 - 현재 입력 필드 이후의 모든 필드에 새 날짜 적용"""
        content = entry.get()
        if len(content) >= 8:
            new_date = content[:8]
            if new_date != self.current_date:
                self.current_date = new_date
                # 현재 필드의 인덱스 찾기
                current_index = self.all_datetime_entries.index(entry)
                # 현재 필드 이후의 모든 필드 업데이트
                for e in self.all_datetime_entries[current_index + 1:]:
                    current_content = e.get()
                    if not current_content:
                        # 비어있는 필드는 새 날짜만 입력
                        e.insert(0, self.current_date + " ")
                    elif len(current_content) >= 8:
                        # 이미 값이 있는 필드는 날짜 부분만 교체
                        time_part = current_content[8:] if len(current_content) > 8 else " "
                        e.delete(0, 'end')
                        e.insert(0, self.current_date + time_part)

    def setup_arrival_tab(self):
        # Timeline 섹션
        timeline_frame = ttk.LabelFrame(self.arrival_tab, text="Timeline")
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
        draft_frame = ttk.LabelFrame(self.arrival_tab, text="Draft")
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
        pilots_frame = ttk.LabelFrame(self.arrival_tab, text="Pilots")
        pilots_frame.pack(pady=5, fill='x', padx=10)
        
        pilot_fields = [("From", "arr_pilot_from_entry"), ("To", "arr_pilot_to_entry")]
        for label_text, entry_name in pilot_fields:
            self.create_datetime_entry(pilots_frame, label_text, entry_name)

        # Towages 섹션
        towages_frame = ttk.LabelFrame(self.arrival_tab, text="Towages")
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

    def setup_operations_tab(self):
        # Gang Operation 섹션 추가
        gang_operation_frame = ttk.LabelFrame(self.operations_tab, text="Gang Operation")
        gang_operation_frame.pack(pady=5, fill='x', padx=70)

        # Operations Start - 먼저 생성
        frame = ttk.Frame(gang_operation_frame)
        frame.pack(pady=5, fill='x')
        self.operationsstart_entry = self.create_datetime_entry(frame, "Operations Start")
        self.operationsstart_entry.pack(side='left', padx=40)

        # Gang Start - datetime entry 생성 시 자동으로 all_datetime_entries에 추가됨
        frame = ttk.Frame(gang_operation_frame)
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
        frame = ttk.Frame(gang_operation_frame)
        frame.pack(pady=5, fill='x')
        self.gangfinish_entry = self.create_datetime_entry(frame, "Gang Finish")
        self.gangfinish_entry.pack(side='left', padx=40)

        # Operations Finish
        frame = ttk.Frame(gang_operation_frame)
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
        frame = ttk.Frame(gang_operation_frame)
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
            # ("BOAAGEONBOA", "boaageonboa_entry"),
            # ("GANORDFOR", "ganordfor_entry"),
            ("Gangway Down", "ganwaydown_entry"),
            ("Operations Commence", "opecom_entry"),
            # ("ESSTTIMOPSCOM", "essttimopscom_entry"), 
            # ("ENTCLECUS", "entclecus_entry"),
            ("Operations Completed", "opecomp_entry"),
            ("Lashing Completed", "lascomp_entry")
        ]
        
        for label_text, entry_name in time_fields:
            self.create_datetime_entry(gang_operation_frame, label_text, entry_name)
            
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
        
        # Operations Completed 값이 변경될 때마다 Lashing Completed 값을 10분 후로 설정
        def update_lashing_completed(*args):
            try:
                # 입력된 값 가져오기
                opecomp_str = self.opecomp_entry.get()
                if not opecomp_str:
                    return
                
                # 입력 형식에 따라 파싱 (YYYYMMDD HHMM 또는 YYYY-MM-DDThh:mm:ss)
                if len(opecomp_str) == 13 and ' ' in opecomp_str:  # YYYYMMDD HHMM 형식
                    date_part = opecomp_str[:8]
                    time_part = opecomp_str[9:]
                    year = date_part[:4]
                    month = date_part[4:6]
                    day = date_part[6:8]
                    hour = time_part[:2]
                    minute = time_part[2:4]
                    opecomp = datetime(int(year), int(month), int(day), int(hour), int(minute))
                else:  # ISO 형식
                    opecomp = datetime.strptime(opecomp_str, "%Y-%m-%dT%H:%M:%S")
                
                # 10분 후 시간 계산
                lashing_completed = opecomp + timedelta(minutes=10)
                
                # lascomp_entry에 설정
                self.lascomp_entry.delete(0, tk.END)
                self.lascomp_entry.insert(0, lashing_completed.strftime("%Y%m%d %H%M"))
                
                print(f"Updated lashing completed to: {lashing_completed.strftime('%Y%m%d %H%M')}")
            except Exception as e:
                print(f"Error updating lashing completed: {str(e)}")
                
        # 이벤트 바인딩 - KeyRelease와 FocusOut 모두 바인딩
        self.opecomp_entry.bind('<KeyRelease>', update_lashing_completed)
        self.opecomp_entry.bind('<FocusOut>', update_lashing_completed)

        # LASCOMPBY 라디오버튼
        frame = ttk.Frame(gang_operation_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="Lashing Done By", width=35).pack(side='left')
        self.lascompby_var = tk.StringVar(value="Terminal")
        
        radio_frame = ttk.Frame(frame)
        radio_frame.pack(side='left')
        ttk.Radiobutton(radio_frame, text="Terminal", variable=self.lascompby_var, 
                       value="Terminal").pack(side='left', padx=5)
        ttk.Radiobutton(radio_frame, text="Vessel Crew", variable=self.lascompby_var,
                       value="Vessel Crew").pack(side='left', padx=5)

    def add_gang_operation_line(self):
        frame = ttk.Frame(self.gang_operation_lines_frame)
        frame.pack(pady=5, fill='x')
        
        # Count
        count_var = tk.StringVar(value="1")
        ttk.Entry(frame, textvariable=count_var, width=6).pack(side='left', padx=2)
        
        # Start Time
        start_time_entry = ttk.Entry(frame, width=15)
        start_time_entry.pack(side='left', padx=2)
        
        # End Time
        end_time_entry = ttk.Entry(frame, width=15)
        end_time_entry.pack(side='left', padx=2)
        
        # 삭제 버튼
        delete_button = ttk.Button(frame, text="X",
                                 command=lambda: self.delete_gang_operation_line(frame))
        delete_button.pack(side='left', padx=2)
        
        line_data = {
            "frame": frame,
            "count": count_var,
            "start_time": start_time_entry,
            "end_time": end_time_entry
        }
        
        self.gang_operation_lines.append(line_data)

    def delete_gang_operation_line(self, frame):
        frame.destroy()
        self.gang_operation_lines = [line for line in self.gang_operation_lines 
                                   if line["frame"] != frame]

    def handle_gang_operation_paste(self):
        try:
            # 클립보드에서 데이터 가져오기
            clipboard_data = self.root.clipboard_get()
            if not clipboard_data.strip():
                messagebox.showwarning("Warning", "클립보드가 비어있습니다.")
                return
            
            # QC 작업 시간 데이터 파싱
            qc_times = []
            lines = [line.strip() for line in clipboard_data.strip().split('\n')]
            
            for line in lines:
                try:
                    parts = [part.strip() for part in line.split('\t') if part.strip()]
                    if len(parts) >= 3 and parts[0].startswith('QC'):
                        start_str = parts[1].strip()
                        end_str = parts[2].strip()
                        
                        try:
                            start_time = datetime.strptime(start_str, "%b-%d-%Y %H:%M")
                            end_time = datetime.strptime(end_str, "%b-%d-%Y %H:%M")
                            qc_times.append((start_time, end_time))
                        except ValueError as ve:
                            print(f"날짜 파싱 오류: {ve}")
                            continue
                except Exception as e:
                    print(f"라인 파싱 오류: {line}, 에러: {str(e)}")
                    continue

            if not qc_times:
                messagebox.showwarning("Warning", "파싱할 수 있는 QC 데이터가 없습니다.")
                return

            # 모든 시간 포인트 수집 (시작 시간과 종료 시간)
            all_times = []
            for start, end in qc_times:
                all_times.append(start)
                all_times.append(end)
            
            # 시간 순으로 정렬
            all_times = sorted(list(set(all_times)))

            # 각 시간 구간별 작동 중인 QC 수 계산
            result_data = []
            for i in range(len(all_times) - 1):
                current_time = all_times[i]
                next_time = all_times[i + 1]
                
                # 현재 시간대에 작동 중인 QC 수 계산
                active_qcs = sum(1 for start, end in qc_times 
                               if start <= current_time and end > current_time)
                
                if active_qcs > 0:
                    # 시작 시간에 1분 추가
                    adjusted_start = current_time + timedelta(minutes=1)
                    result_data.append((active_qcs, adjusted_start, next_time))

            # 기존 라인 삭제
            for line in self.gang_operation_lines:
                line["frame"].destroy()
            self.gang_operation_lines.clear()

            # 새 데이터로 라인 생성
            for count, start, end in result_data:
                frame = ttk.Frame(self.gang_operation_lines_frame)
                frame.pack(pady=5, fill='x')
                
                # Count
                count_var = tk.StringVar(value=str(count))
                ttk.Entry(frame, textvariable=count_var, width=6).pack(side='left', padx=2)
                
                # Start Time
                start_time_entry = ttk.Entry(frame, width=15)
                start_time_entry.insert(0, start.strftime("%Y%m%d %H%M"))
                start_time_entry.pack(side='left', padx=2)
                
                # End Time
                end_time_entry = ttk.Entry(frame, width=15)
                end_time_entry.insert(0, end.strftime("%Y%m%d %H%M"))
                end_time_entry.pack(side='left', padx=2)
                
                # 삭제 버튼
                delete_button = ttk.Button(frame, text="X",
                                         command=lambda f=frame: self.delete_gang_operation_line(f))
                delete_button.pack(side='left', padx=2)
                
                line_data = {
                    "frame": frame,
                    "count": count_var,
                    "start_time": start_time_entry,
                    "end_time": end_time_entry
                }
                
                self.gang_operation_lines.append(line_data)

            messagebox.showinfo("Success", f"{len(result_data)}개의 QC 작업 데이터가 추가되었습니다.")

        except Exception as e:
            error_msg = f"QC 데이터 처리 중 오류 발생: {str(e)}"
            messagebox.showerror("Error", error_msg)

    def setup_departure_tab(self):
        # Timeline 섹션
        timeline_frame = ttk.LabelFrame(self.departure_tab, text="Timeline")
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

        # 날짜/시간 입력 필드들 - create_datetime_entry 사용
        # frame = ttk.Frame(timeline_frame)
        # frame.pack(pady=5, fill='x')
        # self.deppiltugordfor_entry = self.create_datetime_entry(frame, "DEPPILTUGORDFOR")
        # self.deppiltugordfor_entry.pack(side='left', padx=40)

        # frame = ttk.Frame(timeline_frame)
        # frame.pack(pady=5, fill='x')
        # self.boaageoffves_entry = self.create_datetime_entry(frame, "BOAAGEOFFVES")
        # self.boaageoffves_entry.pack(side='left', padx=40)

        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        self.vesundoc_entry = self.create_datetime_entry(frame, "Sailed From Berth")
        self.vesundoc_entry.pack(side='left', padx=40)

        # frame = ttk.Frame(timeline_frame)
        # frame.pack(pady=5, fill='x')
        # self.vessaifrothipor_entry = self.create_datetime_entry(frame, "VESSAIFROTHIPOR")
        # self.vessaifrothipor_entry.pack(side='left', padx=40)

        # Draft 섹션
        draft_frame = ttk.LabelFrame(self.departure_tab, text="Draft")
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
        pilots_frame = ttk.LabelFrame(self.departure_tab, text="Pilots")
        pilots_frame.pack(pady=5, fill='x', padx=10)
        
        pilot_fields = [("From", "dep_pilot_from_entry"), ("To", "dep_pilot_to_entry")]
        for label_text, entry_name in pilot_fields:
            self.create_datetime_entry(pilots_frame, label_text, entry_name)

        # Towages 섹션
        towages_frame = ttk.LabelFrame(self.departure_tab, text="Towages")
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

    def setup_discharge_tab(self):
        # 컨테이너 라인 프레임
        self.discharge_lines_frame = ttk.Frame(self.discharge_tab)
        self.discharge_lines_frame.pack(pady=5, fill='x', padx=10)
        
        # 초기 컨테이너 라인은 생성하지 않음
        self.discharge_lines = []

    def setup_load_tab(self):
        # 컨테이너 라인 프레임
        self.load_lines_frame = ttk.Frame(self.load_tab)
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

    def setup_shifting_tab(self):
        # Lid Moves 섹션
        lid_frame = ttk.LabelFrame(self.shifting_tab, text="Lid Moves")
        lid_frame.pack(pady=5, fill='x', padx=10)
        
        lid_fields = [("On", "lid_on_entry"), ("Off", "lid_off_entry")]
        for label_text, entry_name in lid_fields:
            frame = ttk.Frame(lid_frame)
            frame.pack(pady=5, fill='x')
            ttk.Label(frame, text=label_text, width=15).pack(side='left')
            entry = ttk.Entry(frame, width=3)  # 2자리 정수값
            entry.pack(side='left')
            setattr(self, entry_name, entry)

        add_button = ttk.Button(self.shifting_tab, text="Add Shifting Line",
                              command=self.add_shifting_line)
        add_button.pack(pady=5)
    
        # Container Shifting 섹션
        self.shifting_lines_frame = ttk.LabelFrame(self.shifting_tab, text="Container Shifting")
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
            # XML 생성기 초기화
            generator = VCIXMLGenerator()
            
            # Header 정보 추가
            generator.add_header(
                vessel=self.vessel_entry.get(),
                voyage=self.voyage_entry.get(),
                portun=self.portun_entry.get(),
                master=".",  # 고정값
                berth="1"    # 고정값
            )
            
            # Arrival 정보 추가
            arrival = SubElement(generator.root, 'arrival')
            
            # Arrival Timeline
            timeline = SubElement(arrival, 'timeline')
            self.add_timeline_field(timeline, 'BOWTHARR', 'S', self.bowtharr_var.get())
            # self.add_timeline_field(timeline, 'PILNOT', 'D', 
            #     self.convert_datetime(self.pilot_time_entry.get()))
            # self.add_timeline_field(timeline, 'PILORDFOR', 'D', 
            #     self.convert_datetime(self.pilord_time_entry.get()))
            self.add_timeline_field(timeline, 'ARRPILSTA', 'D', 
                self.convert_datetime(self.arrpil_time_entry.get()))
            # self.add_timeline_field(timeline, 'FIRLINASH', 'D', 
            #     self.convert_datetime(self.firline_time_entry.get()))
            self.add_timeline_field(timeline, 'REAFORANC', 'S', self.arr_reaforanc_var.get())
            
            # Arrival Draft
            draft = SubElement(arrival, 'draft')
            self.add_draft_item(draft, 'AFT', self.arr_draft_aft_entry.get())
            self.add_draft_item(draft, 'FWD', self.arr_draft_fwd_entry.get())
            
            # Arrival Pilots
            pilots = SubElement(arrival, 'pilots')
            pilots.set('cancelled', 'false')
            pilot = SubElement(pilots, 'pilot')
            pilot.set('type', 'Sea')
            pilot.set('number', '1')
            pilot.set('from', self.convert_datetime(self.arr_pilot_from_entry.get()))
            pilot.set('to', self.convert_datetime(self.arr_pilot_to_entry.get()))
            
            # Arrival Towages
            towages = SubElement(arrival, 'towages')
            towages.set('cancelled', 'false')
            
            # 두 예인선 정보 추가
            # tug_names = ["VB ANTARES", "VB LUSITANIA"]
            for i, entries in enumerate(self.arr_tug_entries):
                tug = SubElement(towages, 'tug')
                tug.set('type', 'Sea')
                tug.set('number', '1')
                tug.set('from', self.convert_datetime(entries['From'].get()))
                tug.set('to', self.convert_datetime(entries['To'].get()))
                tug.set('comment', ' ')
                tug.set('tugtype', 'Conventional')
                tug.set('name', ' ')
                tug.set('bowthrusternonop', '0')
            
            # Operations 정보 추가
            operations = SubElement(generator.root, 'operations')
            
            # Operations Timeline
            timeline = SubElement(operations, 'timeline')
            self.add_timeline_field(timeline, 'DOCSIDTO', 'S', self.docsidto_var.get())
            self.add_timeline_field(timeline, 'DOCATTER', 'D', 
                self.convert_datetime(self.docatter_entry.get()))
            # self.add_timeline_field(timeline, 'BOAAGEONBOA', 'D', 
            #     self.convert_datetime(self.boaageonboa_entry.get()))
            # self.add_timeline_field(timeline, 'GANORDFOR', 'D', 
            #     self.convert_datetime(self.ganordfor_entry.get()))
            self.add_timeline_field(timeline, 'GANWAYDOWN', 'D', 
                self.convert_datetime(self.ganwaydown_entry.get()))
            self.add_timeline_field(timeline, 'OPECOM', 'D', 
                self.convert_datetime(self.opecom_entry.get()))
            # self.add_timeline_field(timeline, 'ESSTTIMOPSCOM', 'D', 
            #     self.convert_datetime(self.essttimopscom_entry.get()))
            # self.add_timeline_field(timeline, 'ENTCLECUS', 'D', 
            #     self.convert_datetime(self.entclecus_entry.get()))
            self.add_timeline_field(timeline, 'OPECOMP', 'D', 
                self.convert_datetime(self.opecomp_entry.get()))
            self.add_timeline_field(timeline, 'LASCOMP', 'D', 
                self.convert_datetime(self.lascomp_entry.get()))
            self.add_timeline_field(timeline, 'LASCOMPBY', 'S', self.lascompby_var.get())
            
            # Departure 정보 추가
            departure = SubElement(generator.root, 'departure')
            
            # Departure Timeline
            timeline = SubElement(departure, 'timeline')
            self.add_timeline_field(timeline, 'BOWTHDEP', 'S', self.bowthdep_var.get())
            # self.add_timeline_field(timeline, 'DEPPILTUGORDFOR', 'D', 
            #     self.convert_datetime(self.deppiltugordfor_entry.get()))
            # self.add_timeline_field(timeline, 'BOAAGEOFFVES', 'D', 
            #     self.convert_datetime(self.boaageoffves_entry.get()))
            self.add_timeline_field(timeline, 'VESUNDOC', 'D', 
                self.convert_datetime(self.vesundoc_entry.get()))
            # self.add_timeline_field(timeline, 'VESSAIFROTHIPOR', 'D', 
            #     self.convert_datetime(self.vessaifrothipor_entry.get()))
            
            # Departure Draft
            draft = SubElement(departure, 'draft')
            self.add_draft_item(draft, 'AFT', self.dep_draft_aft_entry.get())
            self.add_draft_item(draft, 'FWD', self.dep_draft_fwd_entry.get())
            
            # Departure Pilots
            pilots = SubElement(departure, 'pilots')
            pilots.set('cancelled', 'false')
            pilot = SubElement(pilots, 'pilot')
            pilot.set('type', 'Sea')
            pilot.set('number', '1')
            pilot.set('from', self.convert_datetime(self.dep_pilot_from_entry.get()))
            pilot.set('to', self.convert_datetime(self.dep_pilot_to_entry.get()))
            
            # Departure Towages
            towages = SubElement(departure, 'towages')
            towages.set('cancelled', 'false')
            
            # 두 예인선 정보 추가
            for i, entries in enumerate(self.dep_tug_entries):
                tug = SubElement(towages, 'tug')
                tug.set('type', 'Sea')
                tug.set('number', '1')
                tug.set('from', self.convert_datetime(entries['From'].get()))
                tug.set('to', self.convert_datetime(entries['To'].get()))
                tug.set('comment', ' ')
                tug.set('tugtype', 'Conventional')
                tug.set('name', ' ')
                tug.set('bowthrusternonop', '0')
            
            # Discharge Details
            discharge = SubElement(generator.root, 'dischargedetails')
            for line in self.discharge_lines:
                linecode = SubElement(discharge, 'linecode')
                linecode.set('type', line['type'].get())
                linecode.set('operator', line['operator'].get())
                linecode.set('containersize', line['size'].get())
                linecode.set('fullempty', line['fe'].get())
                linecode.set('number', line['number'].get())
                linecode.set('terminal', f"{self.portun_entry.get()}PS")
            
            # Load Details
            load = SubElement(generator.root, 'loaddetails')
            for line in self.load_lines:
                linecode = SubElement(load, 'linecode')
                linecode.set('type', line['type'].get())
                linecode.set('operator', line['operator'].get())
                linecode.set('containersize', line['size'].get())
                linecode.set('fullempty', line['fe'].get())
                linecode.set('number', line['number'].get())
                linecode.set('terminal', f"{self.portun_entry.get()}PS")
            
            # Lid Moves
            lidmoves = SubElement(generator.root, 'lidmoves')
            lid = SubElement(lidmoves, 'lid')
            lid.set('terminal', f"{self.portun_entry.get()}PS")
            lid.set('on', self.lid_on_entry.get())
            lid.set('off', self.lid_off_entry.get())
            
            # Container Shifting
            shifting = SubElement(generator.root, 'containershifting')
            for line in self.shifting_lines:
                shift = SubElement(shifting, 'shift')
                shift.set('account', line['account'].get())
                shift.set('type', line['type'].get())
                shift.set('containersize', line['size'].get())
                shift.set('reason', line['reason'].get())
                shift.set('value', line['value'].get())
                shift.set('fullempty', line['fe'].get())
                shift.set('oog', 'true' if line['oog'].get() else 'false')
                shift.set('reefer', 'true' if line['reefer'].get() else 'false')
                shift.set('imo', 'true' if line['imo'].get() else 'false')
                shift.set('notformscaccount', 'true' if line['notformscaccount'].get() else 'false')
            
            # Husbandry
            husbandry = SubElement(generator.root, 'husbandry')
            # Husbandry 항목들은 현재 GUI에 구현되어 있지 않으므로 기본값으로 설정
            husbandry_items = [
                ('B', 'CASHMAST', '1', '20000.00'),
                ('B', 'CONSVISAFE', '1', '1.00'),
                ('B', 'SANITINSP', '1', '500.00'),
                ('B', 'CUSUSEFEE', '1', '1.00')
            ]
            for item_type, code, value, cost in husbandry_items:
                item = SubElement(husbandry, 'item')
                item.set('type', item_type)
                item.set('code', code)
                item.set('value', value)
                item.set('cost', cost)
            
            # Summary
            summary = SubElement(generator.root, 'summary')
            gang_summary = SubElement(summary, 'gangsummary')
            
            gangstart = SubElement(gang_summary, 'gangstart')
            gangstart.text = self.convert_to_iso_datetime(self.gangstart_entry.get())
            
            gangfinish = SubElement(gang_summary, 'gangfinish')
            gangfinish.text = self.convert_to_iso_datetime(self.gangfinish_entry.get())
            
            operationsstart = SubElement(gang_summary, 'operationsstart')
            operationsstart.text = self.convert_to_iso_datetime(self.operationsstart_entry.get())
            
            operationsfinish = SubElement(gang_summary, 'operationsfinish')
            operationsfinish.text = self.convert_to_iso_datetime(self.operationsfinish_entry.get())
            
            # Terminal Efficiency
            terminal_efficiency = SubElement(generator.root, 'terminalefficiency')
            # Terminal Efficiency 항목들은 현재 GUI에 구현되어 있지 않으므로 기본값으로 설정
            efficiency_items = [
                ('Planner_Time_Msc', '2025-03-12T16:25:00'),
                ('Planner_Time_LP', '2025-03-12T16:25:00'),
                ('Terminal_Time', '2025-03-14T14:26:00'),
                ('CI_Proforma', '2.3'),
                ('CI_Planner', '2.7'),
                ('CI_Gang_Availability', '2.8'),
                ('CI_Sub_Optimal', 'Crane intensity was optimal'),
                ('Starting_Operations', 'No changes made'),
                ('Quantity_Of_Containers', '0'),
                ('Changes_After_Ingate', '0'),
                ('Moved_from_Another_Terminal', '0'),
                ('Qty_of_Containers_before_Arrival', '0'),
                ('Live_Connections', '0'),
                ('Reason_for_Live_Connections', 'No live connections'),
                ('Affect_by_Cut_And_Run', '0'),
                ('Reason_For_Cut_and_Run', 'No Cut and run'),
                ('Average_Yard_Utilisation', '93'),
                ('Remarks', 'Departure idle time ( weather )')
            ]
            for name, value in efficiency_items:
                item = SubElement(terminal_efficiency, name)
                item.text = value
            
            # XML 파일 생성
            tree = ElementTree(generator.root)
            filename = f"{generator.root.find('header/vessel').text}_{generator.root.find('header/voyage').text}_{generator.root.find('header/portun').text}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xml"
            tree.write(filename, encoding='utf-8', xml_declaration=True)
            
            messagebox.showinfo("Success", "XML file generated successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate XML: {str(e)}")

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

    def handle_qc_paste(self):
        try:
            # 클립보드에서 데이터 가져오기
            clipboard_data = self.root.clipboard_get()
            if not clipboard_data.strip():
                messagebox.showwarning("Warning", "클립보드가 비어있습니다.")
                return
            
            # QC 작업 시간 데이터 파싱
            qc_times = []
            lines = [line.strip() for line in clipboard_data.strip().split('\n')]
            
            for line in lines:
                try:
                    parts = [part.strip() for part in line.split('\t') if part.strip()]
                    if len(parts) >= 3 and parts[0].startswith('QC'):
                        start_str = parts[1].strip()
                        end_str = parts[2].strip()
                        
                        try:
                            start_time = datetime.strptime(start_str, "%b-%d-%Y %H:%M")
                            end_time = datetime.strptime(end_str, "%b-%d-%Y %H:%M")
                            qc_times.append((start_time, end_time))
                        except ValueError as ve:
                            print(f"날짜 파싱 오류: {ve}")
                            continue
                except Exception as e:
                    print(f"라인 파싱 오류: {line}, 에러: {str(e)}")
                    continue

            if not qc_times:
                messagebox.showwarning("Warning", "파싱할 수 있는 QC 데이터가 없습니다.")
                return

            # 모든 시간 포인트 수집 (시작 시간과 종료 시간)
            all_times = []
            for start, end in qc_times:
                all_times.append(start)
                all_times.append(end)
            
            # 시간 순으로 정렬
            all_times = sorted(list(set(all_times)))

            # 각 시간 구간별 작동 중인 QC 수 계산
            result_data = []
            for i in range(len(all_times) - 1):
                current_time = all_times[i]
                next_time = all_times[i + 1]
                
                # 현재 시간대에 작동 중인 QC 수 계산
                active_qcs = sum(1 for start, end in qc_times 
                               if start <= current_time and end > current_time)
                
                if active_qcs > 0:
                    # 시작 시간에 1분 추가
                    adjusted_start = current_time + timedelta(minutes=1)
                    result_data.append((active_qcs, adjusted_start, next_time))

            # 트리뷰 초기화
            for item in self.qc_tree.get_children():
                self.qc_tree.delete(item)

            # 결과 데이터 트리뷰에 추가
            for count, start, end in result_data:
                self.qc_tree.insert('', 'end', values=(
                    str(count),
                    start.strftime("%Y%m%d %H%M"),
                    end.strftime("%Y%m%d %H%M")
                ))

            # 디버깅을 위한 파싱된 데이터 출력
            print("\n생성된 결과:")
            for count, start, end in result_data:
                print(f"QC 수: {count}, 시작: {start}, 종료: {end}")

            messagebox.showinfo("Success", f"{len(result_data)}개의 시간대별 QC 작업 데이터가 처리되었습니다.")

        except Exception as e:
            error_msg = f"QC 데이터 처리 중 오류 발생: {str(e)}"
            print(f"Error details: {str(e)}")
            messagebox.showerror("Error", error_msg)

def main():
    root = tk.Tk()
    app = VCIGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
