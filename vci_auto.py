import tkinter as tk
from tkinter import ttk, messagebox
from xml.etree.ElementTree import Element, SubElement, ElementTree
import datetime

# pyinstaller --onefile --noconsole vci_auto.py   

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
        self.root.minsize(800, 600)
        
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
        ttk.Label(frame, text=label_text, width=15).pack(side='left')
        
        # 입력 필드
        entry = ttk.Entry(frame, width=13)
        entry.pack(side='left')
        
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
        ttk.Label(frame, text="BOWTHARR", width=15).pack(side='left')
        self.bowtharr_var = tk.StringVar(value="arrival in order")
        ttk.Combobox(frame, textvariable=self.bowtharr_var, 
                    values=["arrival in order"], state="readonly", width=20).pack(side='left')

        # 날짜/시간 입력 필드들
        time_fields = [
            ("PILNOT", "pilot_time_entry"),
            ("PILORDFOR", "pilord_time_entry"),
            ("ARRPILSTA", "arrpil_time_entry"),
            ("FIRLINASH", "firline_time_entry")
        ]
        
        for label_text, entry_name in time_fields:
            self.create_datetime_entry(timeline_frame, label_text, entry_name)

        # REAFORANC 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="REAFORANC", width=20).pack(side='left')
        self.reaforanc_var = tk.StringVar(value="Less than 1 hour")
        ttk.Combobox(frame, textvariable=self.reaforanc_var,
                    values=["Less than 1 hour"], state="readonly").pack(side='left')

        # Draft 섹션
        draft_frame = ttk.LabelFrame(self.arrival_tab, text="Draft")
        draft_frame.pack(pady=5, fill='x', padx=10)
        
        draft_fields = [("AFT", "arr_draft_aft_entry"), ("FWD", "arr_draft_fwd_entry")]
        for label_text, entry_name in draft_fields:
            frame = ttk.Frame(draft_frame)
            frame.pack(pady=5, fill='x')
            ttk.Label(frame, text=label_text, width=15).pack(side='left')
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
        towages_frame.pack(pady=5, fill='x', padx=10)
        
        self.arr_tug_entries = []
        for i in range(2):
            tug_frame = ttk.LabelFrame(towages_frame, text=f"Tug {i+1}")
            tug_frame.pack(pady=5, fill='x')
            
            entries = {}
            # From/To 필드는 datetime entry로 생성
            self.create_datetime_entry(tug_frame, "From", f"arr_tug{i+1}_from_entry")
            self.create_datetime_entry(tug_frame, "To", f"arr_tug{i+1}_to_entry")
            
            # Comment 필드는 일반 entry로 생성
            comment_frame = ttk.Frame(tug_frame)
            comment_frame.pack(pady=5, fill='x')
            ttk.Label(comment_frame, text="Comment", width=15).pack(side='left')
            comment_entry = ttk.Entry(comment_frame, width=20)
            comment_entry.pack(side='left')
            setattr(self, f"arr_tug{i+1}_comment_entry", comment_entry)
            
            entries["From"] = getattr(self, f"arr_tug{i+1}_from_entry")
            entries["To"] = getattr(self, f"arr_tug{i+1}_to_entry")
            entries["Comment"] = comment_entry
            self.arr_tug_entries.append(entries)

    def setup_operations_tab(self):
        # Timeline 섹션
        timeline_frame = ttk.LabelFrame(self.operations_tab, text="Timeline")
        timeline_frame.pack(pady=5, fill='x', padx=10)
        
        # DOCSIDTO 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="DOCSIDTO", width=20).pack(side='left')
        self.docsidto_var = tk.StringVar(value="Port")
        ttk.Combobox(frame, textvariable=self.docsidto_var,
                    values=["Port"], state="readonly").pack(side='left')

        # 날짜/시간 입력 필드들
        time_fields = [
            ("DOCATTER", "docatter_entry"),
            ("BOAAGEONBOA", "boaageonboa_entry"),
            ("GANORDFOR", "ganordfor_entry"),
            ("GANWAYDOWN", "ganwaydown_entry"),
            ("OPECOM", "opecom_entry"),
            ("ESSTTIMOPSCOM", "essttimopscom_entry"),
            ("ENTCLECUS", "entclecus_entry"),
            ("OPECOMP", "opecomp_entry"),
            ("LASCOMP", "lascomp_entry")
        ]
        
        for label_text, entry_name in time_fields:
            self.create_datetime_entry(timeline_frame, label_text, entry_name)

        # LASCOMPBY 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="LASCOMPBY", width=20).pack(side='left')
        self.lascompby_var = tk.StringVar(value="Terminal")
        ttk.Combobox(frame, textvariable=self.lascompby_var,
                    values=["Terminal"], state="readonly").pack(side='left')

    def setup_departure_tab(self):
        # Timeline 섹션
        timeline_frame = ttk.LabelFrame(self.departure_tab, text="Timeline")
        timeline_frame.pack(pady=5, fill='x', padx=10)
        
        # BOWTHDEP 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="BOWTHDEP", width=20).pack(side='left')
        self.bowthdep_var = tk.StringVar(value="departure in order")
        ttk.Combobox(frame, textvariable=self.bowthdep_var,
                    values=["departure in order"], state="readonly").pack(side='left')

        # 날짜/시간 입력 필드들
        time_fields = [
            ("DEPPILTUGORDFOR", "deppiltugordfor_entry"),
            ("BOAAGEOFFVES", "boaageoffves_entry"),
            ("VESUNDOC", "vesundoc_entry"),
            ("VESSAIFROTHIPOR", "vessaifrothipor_entry")
        ]
        
        for label_text, entry_name in time_fields:
            self.create_datetime_entry(timeline_frame, label_text, entry_name)

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
        towages_frame.pack(pady=5, fill='x', padx=10)
        
        self.dep_tug_entries = []
        for i in range(2):
            tug_frame = ttk.LabelFrame(towages_frame, text=f"Tug {i+1}")
            tug_frame.pack(pady=5, fill='x')
            
            entries = {}
            # From/To 필드는 datetime entry로 생성
            self.create_datetime_entry(tug_frame, "From", f"dep_tug{i+1}_from_entry")
            self.create_datetime_entry(tug_frame, "To", f"dep_tug{i+1}_to_entry")
            
            # Comment 필드는 일반 entry로 생성
            comment_frame = ttk.Frame(tug_frame)
            comment_frame.pack(pady=5, fill='x')
            ttk.Label(comment_frame, text="Comment", width=15).pack(side='left')
            comment_entry = ttk.Entry(comment_frame, width=20)
            comment_entry.pack(side='left')
            setattr(self, f"dep_tug{i+1}_comment_entry", comment_entry)
            
            entries["From"] = getattr(self, f"dep_tug{i+1}_from_entry")
            entries["To"] = getattr(self, f"dep_tug{i+1}_to_entry")
            entries["Comment"] = comment_entry
            self.dep_tug_entries.append(entries)

    def setup_discharge_tab(self):
        # 컨테이너 라인 프레임
        self.discharge_lines_frame = ttk.Frame(self.discharge_tab)
        self.discharge_lines_frame.pack(pady=5, fill='x', padx=10)
        
        # 라인 추가 버튼
        add_button = ttk.Button(self.discharge_tab, text="Add Line",
                              command=lambda: self.add_container_line("discharge"))
        add_button.pack(pady=5)
        
        self.discharge_lines = []
        self.add_container_line("discharge")  # 초기 라인 하나 추가

    def setup_load_tab(self):
        # 컨테이너 라인 프레임
        self.load_lines_frame = ttk.Frame(self.load_tab)
        self.load_lines_frame.pack(pady=5, fill='x', padx=10)
        
        # 라인 추가 버튼
        add_button = ttk.Button(self.load_tab, text="Add Line",
                              command=lambda: self.add_container_line("load"))
        add_button.pack(pady=5)
        
        self.load_lines = []
        self.add_container_line("load")  # 초기 라인 하나 추가

    def add_container_line(self, tab_type):
        frame = ttk.Frame(self.discharge_lines_frame if tab_type == "discharge" 
                         else self.load_lines_frame)
        frame.pack(pady=5, fill='x')
        
        # 타입 선택
        type_values = ["DIMP", "TRAN"] if tab_type == "discharge" else ["DEXP", "TRAN"]
        type_var = tk.StringVar(value=type_values[0])
        type_combo = ttk.Combobox(frame, textvariable=type_var,
                                 values=type_values, state="readonly", width=8)
        type_combo.pack(side='left', padx=2)
        
        # 기타 필드들
        operator_var = tk.StringVar(value="MSC")
        ttk.Combobox(frame, textvariable=operator_var,
                    values=["MSC"], state="readonly", width=5).pack(side='left', padx=2)
        
        size_var = tk.StringVar(value="40")
        ttk.Combobox(frame, textvariable=size_var,
                    values=["20", "40"], state="readonly", width=4).pack(side='left', padx=2)
        
        fe_var = tk.StringVar(value="F")
        ttk.Combobox(frame, textvariable=fe_var,
                    values=["F", "E"], state="readonly", width=3).pack(side='left', padx=2)
        
        number_entry = ttk.Entry(frame, width=6)  # 컨테이너 수량
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

        # Container Shifting 섹션
        self.shifting_lines_frame = ttk.Frame(self.shifting_tab)
        self.shifting_lines_frame.pack(pady=5, fill='x', padx=10)
        
        add_button = ttk.Button(self.shifting_tab, text="Add Shifting Line",
                              command=self.add_shifting_line)
        add_button.pack(pady=5)
        
        self.shifting_lines = []
        self.add_shifting_line()  # 초기 라인 하나 추가

    def add_shifting_line(self):
        frame = ttk.Frame(self.shifting_lines_frame)
        frame.pack(pady=5, fill='x')
        
        # 1. Account
        account_var = tk.StringVar(value="MSCU")
        ttk.Combobox(frame, textvariable=account_var,
                    values=["MSCU"], state="readonly", width=8).pack(side='left', padx=2)
        
        # 2. Type
        type_var = tk.StringVar(value="Restow")
        ttk.Combobox(frame, textvariable=type_var,
                    values=["Restow"], state="readonly", width=8).pack(side='left', padx=2)
        
        # 3. Container Size
        size_var = tk.StringVar(value="40")
        ttk.Combobox(frame, textvariable=size_var,
                    values=["20", "40"], state="readonly", width=4).pack(side='left', padx=2)
        
        # 4. Value
        value_entry = ttk.Entry(frame, width=4)  # 컨테이너 수량
        value_entry.pack(side='left', padx=2)
        
        # 5. Full/Empty
        fe_var = tk.StringVar(value="F")
        ttk.Combobox(frame, textvariable=fe_var,
                    values=["F", "E"], state="readonly", width=3).pack(side='left', padx=2)
        
        # 6. OOG
        oog_var = tk.StringVar(value="0")
        ttk.Combobox(frame, textvariable=oog_var,
                    values=["0", "1"], state="readonly", width=3).pack(side='left', padx=2)
        
        # 7. Reefer
        reefer_var = tk.StringVar(value="0")
        ttk.Combobox(frame, textvariable=reefer_var,
                    values=["0", "1"], state="readonly", width=3).pack(side='left', padx=2)
        
        # 8. IMO
        imo_var = tk.StringVar(value="0")
        ttk.Combobox(frame, textvariable=imo_var,
                    values=["0", "1"], state="readonly", width=3).pack(side='left', padx=2)
        
        # 9. Reason
        reason_var = tk.StringVar(value="Restow of optional cargo onboard, to maximize vsl capacity")
        ttk.Combobox(frame, textvariable=reason_var,
                    values=["Restow of optional cargo onboard, to maximize vsl capacity"], 
                    state="readonly", width=50).pack(side='left', padx=2)
        
        # 삭제 버튼
        delete_button = ttk.Button(frame, text="X",
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
            "reason": reason_var
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
            self.add_timeline_field(timeline, 'PILNOT', 'D', 
                self.convert_datetime(self.pilot_time_entry.get()))
            self.add_timeline_field(timeline, 'PILORDFOR', 'D', 
                self.convert_datetime(self.pilord_time_entry.get()))
            self.add_timeline_field(timeline, 'ARRPILSTA', 'D', 
                self.convert_datetime(self.arrpil_time_entry.get()))
            self.add_timeline_field(timeline, 'FIRLINASH', 'D', 
                self.convert_datetime(self.firline_time_entry.get()))
            self.add_timeline_field(timeline, 'REAFORANC', 'S', self.reaforanc_var.get())
            
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
            tug_names = ["VB ANTARES", "VB LUSITANIA"]
            for i, entries in enumerate(self.arr_tug_entries):
                tug = SubElement(towages, 'tug')
                tug.set('type', 'Sea')
                tug.set('number', '1')
                tug.set('from', self.convert_datetime(entries['From'].get()))
                tug.set('to', self.convert_datetime(entries['To'].get()))
                tug.set('comment', tug_names[i])
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
            self.add_timeline_field(timeline, 'BOAAGEONBOA', 'D', 
                self.convert_datetime(self.boaageonboa_entry.get()))
            self.add_timeline_field(timeline, 'GANORDFOR', 'D', 
                self.convert_datetime(self.ganordfor_entry.get()))
            self.add_timeline_field(timeline, 'GANWAYDOWN', 'D', 
                self.convert_datetime(self.ganwaydown_entry.get()))
            self.add_timeline_field(timeline, 'OPECOM', 'D', 
                self.convert_datetime(self.opecom_entry.get()))
            self.add_timeline_field(timeline, 'ESSTTIMOPSCOM', 'D', 
                self.convert_datetime(self.essttimopscom_entry.get()))
            self.add_timeline_field(timeline, 'ENTCLECUS', 'D', 
                self.convert_datetime(self.entclecus_entry.get()))
            self.add_timeline_field(timeline, 'OPECOMP', 'D', 
                self.convert_datetime(self.opecomp_entry.get()))
            self.add_timeline_field(timeline, 'LASCOMP', 'D', 
                self.convert_datetime(self.lascomp_entry.get()))
            self.add_timeline_field(timeline, 'LASCOMPBY', 'S', self.lascompby_var.get())
            
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
            
            # Shifting Details
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
                shift.set('oog', line['oog'].get())
                shift.set('reefer', line['reefer'].get())
                shift.set('imo', line['imo'].get())
            
            # Departure 정보 추가
            departure = SubElement(generator.root, 'departure')
            
            # Departure Timeline
            timeline = SubElement(departure, 'timeline')
            self.add_timeline_field(timeline, 'BOWTHDEP', 'S', self.bowthdep_var.get())
            self.add_timeline_field(timeline, 'DEPPILTUGORDFOR', 'D', 
                self.convert_datetime(self.deppiltugordfor_entry.get()))
            self.add_timeline_field(timeline, 'BOAAGEOFFVES', 'D', 
                self.convert_datetime(self.boaageoffves_entry.get()))
            self.add_timeline_field(timeline, 'VESUNDOC', 'D', 
                self.convert_datetime(self.vesundoc_entry.get()))
            self.add_timeline_field(timeline, 'VESSAIFROTHIPOR', 'D', 
                self.convert_datetime(self.vessaifrothipor_entry.get()))
            
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
            tug_names = ["VB ANTARES", "VB LUSITANIA"]
            for i, entries in enumerate(self.dep_tug_entries):
                tug = SubElement(towages, 'tug')
                tug.set('type', 'Sea')
                tug.set('number', '1')
                tug.set('from', self.convert_datetime(entries['From'].get()))
                tug.set('to', self.convert_datetime(entries['To'].get()))
                tug.set('comment', tug_names[i])
                tug.set('tugtype', 'Conventional')
                tug.set('name', ' ')
                tug.set('bowthrusternonop', '0')
            
            # XML 파일 생성
            tree = ElementTree(generator.root)
            tree.write('output.xml', encoding='utf-8', xml_declaration=True)
            
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

def main():
    root = tk.Tk()
    app = VCIGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
