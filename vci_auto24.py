import tkinter as tk
from tkinter import ttk, messagebox
from xml.etree.ElementTree import Element, SubElement, ElementTree
from datetime import datetime, timedelta
import re

# pyinstaller -w -F --add-binary="C:/Users/kod03/AppData/Local/Programs/Python/Python311/tcl/tkdnd2.8;tkdnd2.8" vci_auto23.py

def parse_excel_data(data):
    """
    엑셀 데이터를 파싱하여 DIMP 및 TRAN 타입별, 사이즈별, FULL/EMPTY 별 합계를 계산합니다.
    """
    try:
        lines = data.strip().split('\n')
        results = {
            'DIMP': {'FULL': {'20': 0, '40': 0, '45': 0},
                     'EMPTY': {'20': 0, '40': 0, '45': 0}},
            'TRAN': {'FULL': {'20': 0, '40': 0, '45': 0},
                     'EMPTY': {'20': 0, '40': 0, '45': 0}}
        }

        # DIMP (첫 번째 행) 처리
        if len(lines) > 0:
            cells_dimp = lines[0].strip().split('\t')
            
            # 안전하게 인덱스 접근을 위한 함수
            def safe_get_int(cells, index, default=0):
                if len(cells) > index and cells[index].strip() and cells[index].strip().isdigit():
                    return int(cells[index])
                return default
            
            # 첫 번째 열은 'LOCAL' 텍스트이므로 건너뛰고 두 번째 열부터 처리
            results['DIMP']['FULL']['20'] = safe_get_int(cells_dimp, 2)
            results['DIMP']['FULL']['40'] = safe_get_int(cells_dimp, 3) + safe_get_int(cells_dimp, 4)
            results['DIMP']['FULL']['45'] = safe_get_int(cells_dimp, 5)
            results['DIMP']['EMPTY']['20'] = safe_get_int(cells_dimp, 6)
            results['DIMP']['EMPTY']['40'] = safe_get_int(cells_dimp, 7) + safe_get_int(cells_dimp, 8)
            results['DIMP']['EMPTY']['45'] = safe_get_int(cells_dimp, 9)

        # TRAN (두 번째, 세 번째 행) 처리
        if len(lines) > 1:
            cells_tran_row1 = lines[1].strip().split('\t')
            cells_tran_row2 = lines[2].strip().split('\t') if len(lines) > 2 else []

            # 안전하게 인덱스 접근을 위한 함수
            def safe_get_int(cells, index, default=0):
                if len(cells) > index and cells[index].strip() and cells[index].strip().isdigit():
                    return int(cells[index])
                return default

            # 첫 번째 열은 '자 T/S' 또는 '타 T/S' 텍스트이므로 건너뛰고 두 번째 열부터 처리
            # FULL
            results['TRAN']['FULL']['20'] = safe_get_int(cells_tran_row1, 1) + safe_get_int(cells_tran_row2, 1)
            results['TRAN']['FULL']['40'] = safe_get_int(cells_tran_row1, 2) + safe_get_int(cells_tran_row1, 3) + \
                                          safe_get_int(cells_tran_row2, 2) + safe_get_int(cells_tran_row2, 3)
            results['TRAN']['FULL']['45'] = safe_get_int(cells_tran_row1, 4) + safe_get_int(cells_tran_row2, 4)

            # EMPTY
            results['TRAN']['EMPTY']['20'] = safe_get_int(cells_tran_row1, 5) + safe_get_int(cells_tran_row2, 5)
            results['TRAN']['EMPTY']['40'] = safe_get_int(cells_tran_row1, 6) + safe_get_int(cells_tran_row1, 7) + \
                                           safe_get_int(cells_tran_row2, 6) + safe_get_int(cells_tran_row2, 7)
            results['TRAN']['EMPTY']['45'] = safe_get_int(cells_tran_row1, 8) + safe_get_int(cells_tran_row2, 8)

        # 결과 출력
        print("DIMP 타입별 사이즈 합계:")
        for fe in ['FULL', 'EMPTY']:
            print(f"- {fe}:")
            for size in ['20', '40', '45']:
                print(f"  - SIZE {size}: {results['DIMP'][fe][size]}")

        print("\nTRAN 타입별 사이즈 합계:")
        for fe in ['FULL', 'EMPTY']:
            print(f"- {fe}:")
            for size in ['20', '40', '45']:
                print(f"  - SIZE {size}: {results['TRAN'][fe][size]}")

        return results
    except Exception as e:
        print(f"Error in parse_excel_data_detailed: {e}")
        return {}

class VCIXMLGenerator:
    def __init__(self):
        # 최상위 요소 생성
        self.root = Element('vcidata')
        self.root.set('version', '5')
        self.root.set('revision', '0')

    def add_header(self, vessel, voyage, portun, master, berth, general_remark=None):
        header = SubElement(self.root, 'header')
        SubElement(header, 'vessel').text = vessel
        SubElement(header, 'voyage').text = voyage
        SubElement(header, 'portun').text = portun
        SubElement(header, 'master').text = master
        SubElement(header, 'berth').text = str(berth)
        
        # # General Remark 추가
        # if general_remark and general_remark.strip():
        #     # summary 요소 생성
        #     summary = SubElement(self.root, 'summary')
        #     generalremarks = SubElement(summary, 'generalremarks')
            
        #     # 각 줄을 별도의 generalremark 요소로 처리
        #     for line in general_remark.strip().split('\n'):
        #         if line.strip():  # 빈 줄 제외
        #             generalremark = SubElement(generalremarks, 'generalremark')
        #             generalremark.set('comment', line.strip())

class PlaceholderEntry(ttk.Entry):
    def __init__(self, master=None, placeholder="", **kwargs):
        super().__init__(master, **kwargs)
        self.placeholder = placeholder
        self.placeholder_color = 'grey'
        self.default_fg_color = 'black'
        self.is_placeholder = True
        
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        
        self._show_placeholder()
    
    def _on_focus_in(self, e=None):
        if self.is_placeholder:
            self.delete(0, tk.END)
            self['foreground'] = self.default_fg_color
            self.is_placeholder = False
            
    def _on_focus_out(self, e=None):
        if not self.get():
            self._show_placeholder()
            
    def _show_placeholder(self):
        self.delete(0, tk.END)
        self.insert(0, self.placeholder)
        self['foreground'] = self.placeholder_color
        self.is_placeholder = True
            
    def get(self):
        current_value = super().get()
        if self.is_placeholder:
            return ""
        return current_value

class VCIGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("VCI Generator")
        
        # Initialize variables
        self.all_datetime_entries = []
        self.arr_tug_entries = []
        self.dep_tug_entries = []
        self.sailed_from_berth_entry = None
        self.dep_pilot_from_entry = None
        self.dep_pilot_to_entry = None
        
        # 탭 초기화 순서를 명시적으로 정의
        self.tabs_initialized = False
        
        # Gang Split 계산 여부를 추적하는 플래그 추가
        self.gang_split_calculated = False
        
        # XML root element 생성
        self.xml_root = Element('vcidata')
        self.xml_root.set('version', '5')
        self.xml_root.set('revision', '0')
        
        # 현재 날짜값을 저장할 변수 추가
        self.current_date = ""
        
        # 모든 날짜/시간 입력 필드들을 순서대로 저장할 리스트
        self.all_datetime_entries = []
        
        # Dock [All Fast] At Terminal 엔트리 초기화
        self.docatter_entry = None  # 초기화만 하고 나중에 설정
        
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
        self.terminal_efficiency_tab = ttk.Frame(self.notebook)  # 새로운 Terminal Efficiency 탭 추가
        
        # 탭 추가
        self.notebook.add(self.header_tab, text="Header")
        self.notebook.add(self.arrival_tab, text="Arrival")
        self.notebook.add(self.operations_tab, text="Operations")
        self.notebook.add(self.departure_tab, text="Departure")
        self.notebook.add(self.discharge_tab, text="Discharge")
        self.notebook.add(self.load_tab, text="Load")
        self.notebook.add(self.shifting_tab, text="Shifting")
        self.notebook.add(self.terminal_efficiency_tab, text="Terminal Efficiency")  # 새로운 탭 추가
        
        # 각 탭의 내용 초기화
        self.setup_header_tab()
        self.setup_arrival_tab()
        self.setup_operations_tab()
        self.setup_departure_tab()
        self.setup_discharge_tab()
        self.setup_load_tab()
        self.setup_shifting_tab()
        self.setup_terminal_efficiency_tab()  # 새로운 탭 설정 메서드 호출

        # 클립보드 바인딩 추가
        self.discharge_tab.bind('<Control-v>', lambda e: self.handle_paste('discharge'))
        self.load_tab.bind('<Control-v>', lambda e: self.handle_paste('load'))
        self.terminal_efficiency_tab.bind('<Control-v>', lambda e: self.handle_terminal_efficiency_paste())  # 새로운 탭에 클립보드 바인딩 추가

        # 컨테이너 라인 엔트리 초기화
        self.discharge_entries = []
        self.load_entries = []

        # 타이머 설정 - 모든 탭이 초기화된 후 이벤트 바인딩
        self.root.after(1000, self.check_tabs_initialized)

    def check_tabs_initialized(self):
        """모든 탭이 초기화되었는지 확인하고 이벤트 바인딩"""
        if (hasattr(self, 'sailed_from_berth_entry') and self.sailed_from_berth_entry is not None and
            hasattr(self, 'dep_pilot_from_entry') and self.dep_pilot_from_entry is not None and
            hasattr(self, 'dep_pilot_to_entry') and self.dep_pilot_to_entry is not None and
            hasattr(self, 'dep_tug_entries') and self.dep_tug_entries):
            
            print("All tabs initialized, binding events...")
            self.bind_sailed_from_berth_events()
        else:
            print("Waiting for tabs to initialize...")
            self.root.after(500, self.check_tabs_initialized)

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
            
        # General Remark 입력 필드 추가
        remark_frame = ttk.Frame(self.header_tab)
        remark_frame.pack(pady=5, fill='x', padx=10)
        
        remark_label = ttk.Label(remark_frame, text="General Remark", width=15)
        remark_label.pack(side='left')
        
        # 여러 줄 텍스트 입력을 위한 Text 위젯 사용
        self.general_remark_text = tk.Text(remark_frame, width=40, height=5)
        self.general_remark_text.pack(side='left', fill='x', expand=True)
        
        # 설명 레이블 추가
        # help_label = ttk.Label(self.header_tab, text="각 줄은 별도의 General Remark으로 XML에 포함됩니다.", font=("Arial", 8))
        # help_label.pack(pady=2, padx=10, anchor='w')

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
            ("Vessel Arrives At Pilot Station/Roads", "arrpil_time_entry"),
        ]
        
        for label_text, entry_name in time_fields:
            self.create_datetime_entry(timeline_frame, label_text, entry_name)

        # Arrival 탭의 REAFORANC 드롭다운
        frame = ttk.Frame(timeline_frame)
        frame.pack(pady=5, fill='x')
        ttk.Label(frame, text="Lost Time Waiting For A Berth", width=30).pack(side='left')
        self.arr_reaforanc_var = tk.StringVar(value="Berth congestion - On window")  # arr_ 접두어 추가
        ttk.Combobox(frame, textvariable=self.arr_reaforanc_var,
                    values=["Berth congestion - On window","Less than 1 hour","Berth Congestion - Off window","Bunkering","Harbour traffic","Authorities","Quarantine","Arrived ahead of schedule","To avoid additional pilotage costs","To avoid additional towage costs","To avoid additional terminal costs","To avoid other port costs","Vessel repairs","Preferred berthing","Waiting for Transhipment Cargo","Navigation Restriction","Public Holidays","Balast/debalast","Geneva instructions","Other"], 
                    state="readonly", width=30).pack(side='left', padx=40)

        # Draft 섹션
        draft_frame = ttk.LabelFrame(self.arrival_tab, text="Draft")
        draft_frame.pack(pady=5, fill='x', padx=10)
        
        draft_fields = [("AFT", "arr_draft_aft_entry"), ("FWD", "arr_draft_fwd_entry")]
        for label_text, entry_name in draft_fields:
            frame = ttk.Frame(draft_frame)
            frame.pack(pady=5, fill='x')
            ttk.Label(frame, text=label_text, width=33).pack(side='left')
            entry = ttk.Entry(frame, width=6)  # 흘수값 (예: 11.45)
            entry.pack(side='left', padx=18)
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

        # Docked [All Fast] At Terminal 값이 변경될 때마다 Pilot, Tug, Gangway Down 값을 업데이트
        def update_all_from_docatter(*args):
            try:
                # docatter_entry가 초기화되지 않았으면 함수 종료
                if self.docatter_entry is None:
                    return
                    
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
                
                # 1. Gangway Down 시간 계산 (Dock [All Fast] At Terminal + 30분)
                gangway_down = docatter + timedelta(minutes=30)
                
                # ganwaydown_entry에 설정
                if hasattr(self, 'ganwaydown_entry'):
                    self.ganwaydown_entry.delete(0, tk.END)
                    self.ganwaydown_entry.insert(0, gangway_down.strftime("%Y%m%d %H%M"))
                    print(f"Updated gangway down to: {gangway_down.strftime('%Y%m%d %H%M')}")
                else:
                    print("Error: ganwaydown_entry is not initialized")
                
                # 2. Pilot 시간 계산
                pilot_from = docatter - timedelta(hours=1)  # Dock [All Fast] At Terminal - 1시간
                pilot_to = docatter + timedelta(minutes=30)  # Dock [All Fast] At Terminal + 30분
                
                # Pilot 시간 설정
                if hasattr(self, 'arr_pilot_from_entry'):
                    self.arr_pilot_from_entry.delete(0, tk.END)
                    self.arr_pilot_from_entry.insert(0, pilot_from.strftime("%Y%m%d %H%M"))
                    print(f"Updated Pilot From to: {pilot_from.strftime('%Y%m%d %H%M')}")
                
                if hasattr(self, 'arr_pilot_to_entry'):
                    self.arr_pilot_to_entry.delete(0, tk.END)
                    self.arr_pilot_to_entry.insert(0, pilot_to.strftime("%Y%m%d %H%M"))
                    print(f"Updated Pilot To to: {pilot_to.strftime('%Y%m%d %H%M')}")
                
                # 3. Tug 시간 계산
                tug_from = docatter - timedelta(hours=1)  # Dock [All Fast] At Terminal - 1시간
                tug_to = docatter + timedelta(hours=1)  # Dock [All Fast] At Terminal + 1시간
                
                # Tug 시간 설정
                for tug_entry in self.arr_tug_entries:
                    tug_entry["From"].delete(0, tk.END)
                    tug_entry["From"].insert(0, tug_from.strftime("%Y%m%d %H%M"))
                    
                    tug_entry["To"].delete(0, tk.END)
                    tug_entry["To"].insert(0, tug_to.strftime("%Y%m%d %H%M"))
                
                print(f"Updated Tug From to: {tug_from.strftime('%Y%m%d %H%M')}")
                print(f"Updated Tug To to: {tug_to.strftime('%Y%m%d %H%M')}")
                
            except Exception as e:
                print(f"Error updating fields from Dock [All Fast] At Terminal: {str(e)}")
                import traceback
                traceback.print_exc()
                
        # docatter_entry가 초기화된 후에 이벤트 바인딩
        def bind_docatter_events():
            if self.docatter_entry is not None:
                self.docatter_entry.bind('<KeyRelease>', update_all_from_docatter)
                self.docatter_entry.bind('<FocusOut>', update_all_from_docatter)
                print("Successfully bound events to docatter_entry")
            else:
                # docatter_entry가 아직 초기화되지 않았으면 100ms 후에 다시 시도
                self.root.after(100, bind_docatter_events)
        
        # 바인딩 시도 시작
        bind_docatter_events()

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
        gang_operation_frame.pack(pady=5, fill='x', padx=10)
        
        # Operations Start - 먼저 생성
        frame = ttk.Frame(gang_operation_frame)
        frame.pack(fill='x')
        label = ttk.Label(frame, text="Operations Start", width=30, background='#FFCCCC')
        label.pack(side='left')
        self.operationsstart_entry = ttk.Entry(frame, width=13)
        self.operationsstart_entry.pack(side='left', padx=40)
        self.operationsstart_entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(self.operationsstart_entry))
        self.operationsstart_entry.bind('<KeyRelease>', lambda e: self.handle_date_change(self.operationsstart_entry))
        self.all_datetime_entries.append(self.operationsstart_entry)
        ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)
        self.operationsstart_entry.pack(side='left', fill='x', padx=40)

        # Sailed From Berth - datetime entry 생성
        frame = ttk.Frame(gang_operation_frame)
        frame.pack(fill='x')
        label = ttk.Label(frame, text="Sailed From Berth", width=30, background='#FFCCCC')
        label.pack(side='left')
        self.sailed_from_berth_entry = ttk.Entry(frame, width=13)
        self.sailed_from_berth_entry.pack(side='left', padx=40)
        self.sailed_from_berth_entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(self.sailed_from_berth_entry))
        self.sailed_from_berth_entry.bind('<KeyRelease>', lambda e: self.handle_date_change(self.sailed_from_berth_entry))
        self.all_datetime_entries.append(self.sailed_from_berth_entry)
        ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)

        # Sailed From Berth 값이 변경될 때마다 Departure 탭의 Pilot, Tug 값을 업데이트
        def update_departure_times(*args):
            try:
                # 입력된 값 가져오기
                sailed_str = self.sailed_from_berth_entry.get()
                if not sailed_str:
                    return
                
                print(f"Processing sailed_str: {sailed_str}")
                
                # 입력 형식에 따라 파싱 (YYYYMMDD HHMM 또는 YYYY-MM-DDThh:mm:ss)
                if len(sailed_str) == 13 and ' ' in sailed_str:  # YYYYMMDD HHMM 형식
                    date_part = sailed_str[:8]
                    time_part = sailed_str[9:]
                    year = date_part[:4]
                    month = date_part[4:6]
                    day = date_part[6:8]
                    hour = time_part[:2]
                    minute = time_part[2:4]
                    sailed_time = datetime(int(year), int(month), int(day), int(hour), int(minute))
                else:  # ISO 형식
                    sailed_time = datetime.strptime(sailed_str, "%Y-%m-%dT%H:%M:%S")
                
                print(f"Parsed sailed_time: {sailed_time}")
                
                # 1. Pilot 시간 계산
                pilot_from = sailed_time - timedelta(hours=1)  # Sailed From Berth - 1시간
                pilot_to = sailed_time + timedelta(minutes=30)  # Sailed From Berth + 30분
                
                # Pilot 시간 설정
                if hasattr(self, 'dep_pilot_from_entry') and self.dep_pilot_from_entry is not None:
                    self.dep_pilot_from_entry.delete(0, tk.END)
                    self.dep_pilot_from_entry.insert(0, pilot_from.strftime("%Y%m%d %H%M"))
                    print(f"Updated Departure Pilot From to: {pilot_from.strftime('%Y%m%d %H%M')}")
                else:
                    print("dep_pilot_from_entry is not initialized")
                
                if hasattr(self, 'dep_pilot_to_entry') and self.dep_pilot_to_entry is not None:
                    self.dep_pilot_to_entry.delete(0, tk.END)
                    self.dep_pilot_to_entry.insert(0, pilot_to.strftime("%Y%m%d %H%M"))
                    print(f"Updated Departure Pilot To to: {pilot_to.strftime('%Y%m%d %H%M')}")
                else:
                    print("dep_pilot_to_entry is not initialized")
                
                # 2. Tug 시간 계산
                tug_from = sailed_time - timedelta(hours=1)  # Sailed From Berth - 1시간
                tug_to = sailed_time + timedelta(hours=1)  # Sailed From Berth + 1시간
                
                # Tug 시간 설정
                if hasattr(self, 'dep_tug_entries') and self.dep_tug_entries:
                    for tug_entry in self.dep_tug_entries:
                        tug_entry["From"].delete(0, tk.END)
                        tug_entry["From"].insert(0, tug_from.strftime("%Y%m%d %H%M"))
                        
                        tug_entry["To"].delete(0, tk.END)
                        tug_entry["To"].insert(0, tug_to.strftime("%Y%m%d %H%M"))
                    
                    print(f"Updated Departure Tug From to: {tug_from.strftime('%Y%m%d %H%M')}")
                    print(f"Updated Departure Tug To to: {tug_to.strftime('%Y%m%d %H%M')}")
                else:
                    print("dep_tug_entries is not initialized or empty")
                
            except Exception as e:
                print(f"Error updating departure times: {str(e)}")
                import traceback
                traceback.print_exc()
                
        # sailed_from_berth_entry에 이벤트 바인딩
        self.sailed_from_berth_entry.bind('<KeyRelease>', update_departure_times)
        self.sailed_from_berth_entry.bind('<FocusOut>', update_departure_times)
        print("Bound events to sailed_from_berth_entry")

        # Gang Finish
        frame = ttk.Frame(gang_operation_frame)
        frame.pack(pady=5, fill='x')
        self.gangfinish_entry = self.create_datetime_entry(frame, "Gang Finish")
        self.gangfinish_entry.pack(side='left', padx=40)

        # Operations Finish
        frame = ttk.Frame(gang_operation_frame)
        frame.pack(pady=5, fill='x')
        label = ttk.Label(frame, text="Operations Finish", width=30, background='#FFCCCC')
        label.pack(side='left')
        self.operationsfinish_entry = ttk.Entry(frame, width=13)
        self.operationsfinish_entry.pack(side='left', padx=40)
        self.operationsfinish_entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(self.operationsfinish_entry))
        self.operationsfinish_entry.bind('<KeyRelease>', lambda e: self.handle_date_change(self.operationsfinish_entry))
        self.all_datetime_entries.append(self.operationsfinish_entry)
        ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)
        self.operationsfinish_entry.pack(side='left', fill='x', padx=40)
        
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
            if label_text == "Docked [All Fast] At Terminal":
                frame = ttk.Frame(gang_operation_frame)
                frame.pack(pady=5, fill='x')
                label = ttk.Label(frame, text=label_text, width=30, background='#FFCCCC')
                label.pack(side='left')
                self.docatter_entry = ttk.Entry(frame, width=13)
                self.docatter_entry.pack(side='left', padx=40)
                self.docatter_entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(self.docatter_entry))
                self.docatter_entry.bind('<KeyRelease>', lambda e: self.handle_date_change(self.docatter_entry))
                self.all_datetime_entries.append(self.docatter_entry)
                ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)
            else:
                entry = self.create_datetime_entry(gang_operation_frame, label_text, entry_name)
                # ganwaydown_entry가 생성되었는지 확인하고 디버그 정보 출력
                if entry_name == "ganwaydown_entry":
                    print(f"ganwaydown_entry created: {entry}")
                    self.ganwaydown_entry = entry
            
        # Docked [All Fast] At Terminal 값이 변경될 때마다 Gangway Down 값을 30분 후로 설정
        def update_gangway_down(*args):
            try:
                # ganwaydown_entry가 초기화되었는지 확인
                if not hasattr(self, 'ganwaydown_entry'):
                    print("Error: ganwaydown_entry is not initialized")
                    return
                
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
                import traceback
                traceback.print_exc()
        
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

        # Gang Split 섹션 추가
        gang_split_frame = ttk.LabelFrame(self.operations_tab, text="Gang Split")
        gang_split_frame.pack(pady=5, fill='x', padx=10)

        # 총 move 갯수와 총 crane 갯수를 표시할 프레임 추가
        summary_frame = ttk.Frame(gang_split_frame)
        summary_frame.pack(pady=5, fill='x')
        
        # 총 move 갯수 레이블
        self.total_moves_label = ttk.Label(summary_frame, text="Total Moves: 0")
        self.total_moves_label.pack(side='left', padx=10)
        
        # 총 crane 갯수 레이블
        self.total_cranes_label = ttk.Label(summary_frame, text="Total Cranes: 0")
        self.total_cranes_label.pack(side='left', padx=10)
        # 계산 버튼 추가
        calculate_button = ttk.Button(summary_frame, text="Calculate Gang Split",
                                    command=self.calculate_gang_split)
        calculate_button.pack(side='left', padx=315)

        # Gang Split 라인을 표시할 프레임
        self.gang_split_lines_frame = ttk.Frame(gang_split_frame)
        self.gang_split_lines_frame.pack(pady=5, fill='x')

        # 헤더 추가
        header_frame = ttk.Frame(self.gang_split_lines_frame)
        header_frame.pack(fill='x')
        ttk.Label(header_frame, text="Crane", width=5).pack(side='left', padx=10)
        ttk.Label(header_frame, text="Cntr", width=5).pack(side='left', padx=5)

        # Gang Split 라인을 저장할 리스트
        self.gang_split_lines = []

        # 크레인 라인을 한 줄에 배치
        row_frame = ttk.Frame(self.gang_split_lines_frame)
        row_frame.pack(pady=2, fill='x')
        
        for crane_id in range(1, 9):
            frame = ttk.Frame(row_frame)
            frame.pack(side='left', padx=2, expand=True)
            
            # Crane ID (읽기 전용)
            ttk.Label(frame, text=str(crane_id), width=3).pack(side='left', padx=1)
            
            # Work Time
            work_time_var = tk.StringVar(value="")
            work_time_entry = ttk.Entry(frame, textvariable=work_time_var, width=5)
            work_time_entry.pack(side='left', padx=1)
            
            line_data = {
                "frame": frame,
                "crane_id": crane_id,
                "work_time": work_time_var
            }
            
            self.gang_split_lines.append(line_data)

        # Gang Work 섹션 추가
        gang_work_frame = ttk.LabelFrame(self.operations_tab, text="Gang Work")
        gang_work_frame.pack(pady=5, fill='x', padx=10, anchor='w')
        
        # 프레임 너비 설정
        gang_work_frame.configure(width=600)
        
        # 상단 버튼과 라벨을 위한 프레임
        top_frame = ttk.Frame(gang_work_frame)
        top_frame.pack(pady=5, fill='x', padx=10)
        
        # 붙여넣기 버튼 추가
        paste_button = ttk.Button(top_frame, text="Paste Gang Work Data",
                                command=self.handle_gang_work_paste)
        paste_button.pack(side='left', padx=5)
        
        # 빈 영역 클릭 시 붙여넣기를 위한 레이블 추가
        empty_area = ttk.Label(top_frame, text="Click here to paste data")
        empty_area.pack(side='left', padx=5)
        
        # 빈 영역 클릭 이벤트 바인딩
        empty_area.bind('<Button-1>', lambda e: self.handle_gang_work_paste())
        
        # 구분선 추가
        separator = ttk.Separator(gang_work_frame, orient='horizontal')
        separator.pack(fill='x', padx=10, pady=5)
        
        # Gang Work 라인을 표시할 프레임 (스크롤 가능한 영역)
        container_frame = ttk.Frame(gang_work_frame)
        container_frame.pack(pady=5, fill='both', expand=True, padx=10)
        
        # 스크롤바 추가
        canvas = tk.Canvas(container_frame)
        scrollbar = ttk.Scrollbar(container_frame, orient="vertical", command=canvas.yview)
        self.gang_work_lines_frame = ttk.Frame(canvas)
        
        # 스크롤 가능한 영역 설정
        self.gang_work_lines_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.gang_work_lines_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 스크롤바와 캔버스 배치
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # 헤더 추가
        header_frame = ttk.Frame(self.gang_work_lines_frame)
        header_frame.pack(fill='x')
        ttk.Label(header_frame, text="Count", width=6).pack(side='left', padx=2)
        ttk.Label(header_frame, text="Start Time", width=15).pack(side='left', padx=2)
        ttk.Label(header_frame, text="End Time", width=15).pack(side='left', padx=2)

        # Gang Work 라인을 저장할 리스트
        self.gang_work_lines = []
        
        # 클립보드 바인딩 추가 (프레임 전체에 적용)
        gang_work_frame.bind('<Control-v>', lambda e: self.handle_gang_work_paste())
        empty_area.bind('<Control-v>', lambda e: self.handle_gang_work_paste())

    def calculate_gang_split(self):
        """Gang Split 계산 함수"""
        try:
            # 총 컨테이너 갯수 계산 (discharge + load + lid + shifting)
            total_containers = 0
            
            # Discharge 컨테이너 합계
            for line in self.discharge_lines:
                try:
                    number = int(line['number'].get())
                    total_containers += number
                except (ValueError, TypeError):
                    pass  # 빈 값이나 숫자가 아닌 값 무시
            
            # Load 컨테이너 합계
            for line in self.load_lines:
                try:
                    number = int(line['number'].get())
                    total_containers += number
                except (ValueError, TypeError):
                    pass  # 빈 값이나 숫자가 아닌 값 무시
            
            # Lid 값 합계
            try:
                lid_on = int(self.lid_on_entry.get())
                total_containers += lid_on
            except (ValueError, TypeError):
                pass  # 빈 값이나 숫자가 아닌 값 무시
                
            try:
                lid_off = int(self.lid_off_entry.get())
                total_containers += lid_off
            except (ValueError, TypeError):
                pass  # 빈 값이나 숫자가 아닌 값 무시
            
            # Shifting 값 합계 (Restow는 2배, Shift는 1배)
            for line in self.shifting_lines:
                try:
                    value = int(line['value'].get())
                    type_ = line['type'].get()
                    if type_ == "Restow":
                        total_containers += value * 2
                    else:  # Shift
                        total_containers += value
                except (ValueError, TypeError):
                    pass  # 빈 값이나 숫자가 아닌 값 무시
            
            # 총 크레인 갯수 계산 (Gang Work의 count 중 최대값)
            max_crane_count = 0
            for line in self.gang_work_lines:
                try:
                    count = int(line['count'].get())
                    max_crane_count = max(max_crane_count, count)
                except (ValueError, TypeError):
                    pass  # 빈 값이나 숫자가 아닌 값 무시
            
            # 크레인 갯수가 0이면 기본값 1 설정
            if max_crane_count == 0:
                max_crane_count = 1
            
            # 각 크레인당 컨테이너 수 계산
            containers_per_crane = total_containers // max_crane_count
            remainder = total_containers % max_crane_count
            
            # 결과를 Gang Split 라인에 입력
            for i, line in enumerate(self.gang_split_lines):
                if i < max_crane_count:
                    # 나머지가 있으면 앞쪽 크레인부터 1개씩 추가
                    if i < remainder:
                        line['work_time'].set(str(containers_per_crane + 1))
                    else:
                        line['work_time'].set(str(containers_per_crane))
                else:
                    # 사용하지 않는 크레인은 빈 값으로 설정
                    line['work_time'].set("")
            
            # 총 move 갯수와 총 crane 갯수 레이블 업데이트
            self.total_moves_label.config(text=f"Total Moves: {total_containers}")
            self.total_cranes_label.config(text=f"Total Cranes: {max_crane_count}")
            
            # Gang Split 계산 완료 플래그 설정
            self.gang_split_calculated = True
            
            messagebox.showinfo("Success", f"Gang Split 계산 완료: 총 {total_containers}개 컨테이너, {max_crane_count}개 크레인")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gang Split 계산 중 오류 발생: {str(e)}")

    def handle_gang_work_paste(self):
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
                    if len(parts) >= 3 and parts[0].startswith('QC'):  # QC번호, 시작시간, 종료시간
                        qc_number = parts[0]  # QC 번호 (예: QC106)
                        start_str = parts[1]  # 시작 시간
                        end_str = parts[2]    # 종료 시간
                        
                        try:
                            # Apr-11-2025 17:30 형식의 날짜를 파싱
                            start_dt = datetime.strptime(start_str, "%b-%d-%Y %H:%M")
                            end_dt = datetime.strptime(end_str, "%b-%d-%Y %H:%M")
                            qc_times.append((start_dt, end_dt))
                        except ValueError as ve:
                            print(f"날짜 파싱 오류: {ve}")
                            continue
                except Exception as e:
                    print(f"라인 파싱 오류: {line}, 에러: {str(e)}")
                    continue

            if not qc_times:
                messagebox.showwarning("Warning", "파싱할 수 있는 Gang Work 데이터가 없습니다.")
                return

            # 모든 시간 포인트 수집 (시작 시간과 종료 시간)
            all_times = []
            for start, end in qc_times:
                all_times.append(start)
                all_times.append(end)
            
            # 시간 순으로 정렬하고 중복 제거
            all_times = sorted(list(set(all_times)))

            # 각 시간 구간별 작동 중인 QC 수 계산
            time_periods = []
            for i in range(len(all_times) - 1):
                current_time = all_times[i]
                next_time = all_times[i + 1]
                
                # 현재 시간대에 작동 중인 QC 수 계산
                active_qcs = sum(1 for start, end in qc_times 
                               if start <= current_time and end > current_time)
                
                if active_qcs > 0:  # QC가 작동 중인 경우만 포함
                    # 첫 번째 구간이 아닌 경우 시작 시간에 1분 추가
                    if i > 0:
                        current_time = current_time + timedelta(minutes=1)
                    time_periods.append((active_qcs, current_time, next_time))

            # 기존 라인 삭제
            for line in self.gang_work_lines:
                line["frame"].destroy()
            self.gang_work_lines.clear()

            # 새 데이터로 라인 생성
            for count, start, end in time_periods:
                frame = ttk.Frame(self.gang_work_lines_frame)
                frame.pack(pady=5, fill='x')
                
                # Count (QC 수)
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
                                         command=lambda f=frame: self.delete_gang_work_line(f))
                delete_button.pack(side='left', padx=2)
                
                line_data = {
                    "frame": frame,
                    "count": count_var,
                    "start_time": start_time_entry,
                    "end_time": end_time_entry
                }
                
                self.gang_work_lines.append(line_data)

            # 디버깅을 위한 파싱된 데이터 출력
            print("\n처리된 Gang Work 데이터:")
            for count, start, end in time_periods:
                print(f"작동 중인 QC 수: {count}, 시작: {start.strftime('%Y-%m-%d %H:%M')}, 종료: {end.strftime('%Y-%m-%d %H:%M')}")

            messagebox.showinfo("Success", f"{len(time_periods)}개의 시간대별 Gang Work 데이터가 추가되었습니다.")

        except Exception as e:
            error_msg = f"Gang Work 데이터 처리 중 오류 발생: {str(e)}"
            print(f"Error details: {str(e)}")  # 디버깅을 위한 상세 에러 출력
            messagebox.showerror("Error", error_msg)

    def delete_gang_work_line(self, frame):
        frame.destroy()
        self.gang_work_lines = [line for line in self.gang_work_lines 
                               if line["frame"] != frame]

    def setup_departure_tab(self):
        # Timeline 섹션
        timeline_frame = ttk.LabelFrame(self.departure_tab, text="Timeline")
        timeline_frame.pack(pady=5, fill='x', padx=10)
        
        # Sailed From Berth - datetime entry 생성
        frame = ttk.Frame(timeline_frame)
        frame.pack(fill='x')
        label = ttk.Label(frame, text="Sailed From Berth", width=30, background='#FFCCCC')
        label.pack(side='left')
        self.sailed_from_berth_entry = ttk.Entry(frame, width=13)
        self.sailed_from_berth_entry.pack(side='left', padx=40)
        self.sailed_from_berth_entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(self.sailed_from_berth_entry))
        self.sailed_from_berth_entry.bind('<KeyRelease>', lambda e: self.handle_date_change(self.sailed_from_berth_entry))
        # 이벤트 바인딩 수정
        self.sailed_from_berth_entry.bind('<KeyRelease>', self.update_departure_times)
        self.sailed_from_berth_entry.bind('<FocusOut>', self.update_departure_times)
        self.all_datetime_entries.append(self.sailed_from_berth_entry)
        ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)
        
        # Pilot From - datetime entry 생성
        frame = ttk.Frame(timeline_frame)
        frame.pack(fill='x')
        label = ttk.Label(frame, text="Pilot From", width=30, background='#FFCCCC')
        label.pack(side='left')
        self.dep_pilot_from_entry = ttk.Entry(frame, width=13)
        self.dep_pilot_from_entry.pack(side='left', padx=40)
        self.dep_pilot_from_entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(self.dep_pilot_from_entry))
        self.dep_pilot_from_entry.bind('<KeyRelease>', lambda e: self.handle_date_change(self.dep_pilot_from_entry))
        self.all_datetime_entries.append(self.dep_pilot_from_entry)
        ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)
        
        # Pilot To - datetime entry 생성
        frame = ttk.Frame(timeline_frame)
        frame.pack(fill='x')
        label = ttk.Label(frame, text="Pilot To", width=30, background='#FFCCCC')
        label.pack(side='left')
        self.dep_pilot_to_entry = ttk.Entry(frame, width=13)
        self.dep_pilot_to_entry.pack(side='left', padx=40)
        self.dep_pilot_to_entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(self.dep_pilot_to_entry))
        self.dep_pilot_to_entry.bind('<KeyRelease>', lambda e: self.handle_date_change(self.dep_pilot_to_entry))
        self.all_datetime_entries.append(self.dep_pilot_to_entry)
        ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)
        
        # Tug 섹션 추가
        tug_frame = ttk.LabelFrame(self.departure_tab, text="Tug")
        tug_frame.pack(pady=5, fill='x', padx=10)
        
        # Tug 버튼 프레임
        tug_button_frame = ttk.Frame(tug_frame)
        tug_button_frame.pack(fill='x', pady=5)
        
        # Tug 추가 버튼
        add_tug_button = ttk.Button(tug_button_frame, text="Add Tug", command=lambda: self.add_tug_frame(tug_frame, "departure"))
        add_tug_button.pack(side='left', padx=5)
        
        # 초기 Tug 프레임 추가
        self.add_tug_frame(tug_frame, "departure")
        
        # 나머지 필드들...

    def bind_sailed_from_berth_events(self):
        """Sailed From Berth 값이 변경될 때마다 Departure 탭의 Pilot, Tug 값을 업데이트하는 이벤트 바인딩"""
        if not hasattr(self, 'sailed_from_berth_entry') or self.sailed_from_berth_entry is None:
            print("sailed_from_berth_entry is not initialized")
            return
            
        # 기존 바인딩 제거 (중복 바인딩 방지)
        try:
            self.sailed_from_berth_entry.unbind('<KeyRelease>')
            self.sailed_from_berth_entry.unbind('<FocusOut>')
        except:
            pass
            
        # 이벤트 핸들러 함수 정의
        def update_departure_times(*args):
            try:
                # 입력된 값 가져오기
                sailed_str = self.sailed_from_berth_entry.get()
                if not sailed_str:
                    return
                
                print(f"Processing sailed_str: {sailed_str}")
                
                # 입력 형식에 따라 파싱 (YYYYMMDD HHMM 또는 YYYY-MM-DDThh:mm:ss)
                if len(sailed_str) == 13 and ' ' in sailed_str:  # YYYYMMDD HHMM 형식
                    date_part = sailed_str[:8]
                    time_part = sailed_str[9:]
                    year = date_part[:4]
                    month = date_part[4:6]
                    day = date_part[6:8]
                    hour = time_part[:2]
                    minute = time_part[2:4]
                    sailed_time = datetime(int(year), int(month), int(day), int(hour), int(minute))
                else:  # ISO 형식
                    sailed_time = datetime.strptime(sailed_str, "%Y-%m-%dT%H:%M:%S")
                
                print(f"Parsed sailed_time: {sailed_time}")
                
                # 1. Pilot 시간 계산
                pilot_from = sailed_time - timedelta(hours=1)  # Sailed From Berth - 1시간
                pilot_to = sailed_time + timedelta(minutes=30)  # Sailed From Berth + 30분
                
                # Pilot 시간 설정
                if hasattr(self, 'dep_pilot_from_entry') and self.dep_pilot_from_entry is not None:
                    self.dep_pilot_from_entry.delete(0, tk.END)
                    self.dep_pilot_from_entry.insert(0, pilot_from.strftime("%Y%m%d %H%M"))
                    print(f"Updated Departure Pilot From to: {pilot_from.strftime('%Y%m%d %H%M')}")
                else:
                    print("dep_pilot_from_entry is not initialized")
                
                if hasattr(self, 'dep_pilot_to_entry') and self.dep_pilot_to_entry is not None:
                    self.dep_pilot_to_entry.delete(0, tk.END)
                    self.dep_pilot_to_entry.insert(0, pilot_to.strftime("%Y%m%d %H%M"))
                    print(f"Updated Departure Pilot To to: {pilot_to.strftime('%Y%m%d %H%M')}")
                else:
                    print("dep_pilot_to_entry is not initialized")
                
                # 2. Tug 시간 계산
                tug_from = sailed_time - timedelta(hours=1)  # Sailed From Berth - 1시간
                tug_to = sailed_time + timedelta(hours=1)  # Sailed From Berth + 1시간
                
                # Tug 시간 설정
                if hasattr(self, 'dep_tug_entries') and self.dep_tug_entries:
                    for tug_entry in self.dep_tug_entries:
                        tug_entry["From"].delete(0, tk.END)
                        tug_entry["From"].insert(0, tug_from.strftime("%Y%m%d %H%M"))
                        
                        tug_entry["To"].delete(0, tk.END)
                        tug_entry["To"].insert(0, tug_to.strftime("%Y%m%d %H%M"))
                    
                    print(f"Updated Departure Tug From to: {tug_from.strftime('%Y%m%d %H%M')}")
                    print(f"Updated Departure Tug To to: {tug_to.strftime('%Y%m%d %H%M')}")
                else:
                    print("dep_tug_entries is not initialized or empty")
                
            except Exception as e:
                print(f"Error updating departure times: {str(e)}")
                import traceback
                traceback.print_exc()
        
        # 이벤트 바인딩
        self.sailed_from_berth_entry.bind('<KeyRelease>', update_departure_times)
        self.sailed_from_berth_entry.bind('<FocusOut>', update_departure_times)
        print("Successfully bound events to sailed_from_berth_entry")
        
        # 수동으로 한 번 실행하여 초기값 설정
        update_departure_times()

        # Gang Finish
        frame = ttk.Frame(gang_operation_frame)
        frame.pack(pady=5, fill='x')
        self.gangfinish_entry = self.create_datetime_entry(frame, "Gang Finish")
        self.gangfinish_entry.pack(side='left', padx=40)

        # Operations Finish
        frame = ttk.Frame(gang_operation_frame)
        frame.pack(pady=5, fill='x')
        label = ttk.Label(frame, text="Operations Finish", width=30, background='#FFCCCC')
        label.pack(side='left')
        self.operationsfinish_entry = ttk.Entry(frame, width=13)
        self.operationsfinish_entry.pack(side='left', padx=40)
        self.operationsfinish_entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(self.operationsfinish_entry))
        self.operationsfinish_entry.bind('<KeyRelease>', lambda e: self.handle_date_change(self.operationsfinish_entry))
        self.all_datetime_entries.append(self.operationsfinish_entry)
        ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)
        self.operationsfinish_entry.pack(side='left', fill='x', padx=40)
        
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
            if label_text == "Docked [All Fast] At Terminal":
                frame = ttk.Frame(gang_operation_frame)
                frame.pack(pady=5, fill='x')
                label = ttk.Label(frame, text=label_text, width=30, background='#FFCCCC')
                label.pack(side='left')
                self.docatter_entry = ttk.Entry(frame, width=13)
                self.docatter_entry.pack(side='left', padx=40)
                self.docatter_entry.bind('<FocusIn>', lambda e: self.handle_entry_focus(self.docatter_entry))
                self.docatter_entry.bind('<KeyRelease>', lambda e: self.handle_date_change(self.docatter_entry))
                self.all_datetime_entries.append(self.docatter_entry)
                ttk.Label(frame, text="Format: YYYYMMDD HHMM").pack(side='left', padx=5)
            else:
                entry = self.create_datetime_entry(gang_operation_frame, label_text, entry_name)
                # ganwaydown_entry가 생성되었는지 확인하고 디버그 정보 출력
                if entry_name == "ganwaydown_entry":
                    print(f"ganwaydown_entry created: {entry}")
                    self.ganwaydown_entry = entry
            
        # Docked [All Fast] At Terminal 값이 변경될 때마다 Gangway Down 값을 30분 후로 설정
        def update_gangway_down(*args):
            try:
                # ganwaydown_entry가 초기화되었는지 확인
                if not hasattr(self, 'ganwaydown_entry'):
                    print("Error: ganwaydown_entry is not initialized")
                    return
                
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
                import traceback
                traceback.print_exc()
        
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

        # Gang Split 섹션 추가
        gang_split_frame = ttk.LabelFrame(self.operations_tab, text="Gang Split")
        gang_split_frame.pack(pady=5, fill='x', padx=10)

        # 총 move 갯수와 총 crane 갯수를 표시할 프레임 추가
        summary_frame = ttk.Frame(gang_split_frame)
        summary_frame.pack(pady=5, fill='x')
        
        # 총 move 갯수 레이블
        self.total_moves_label = ttk.Label(summary_frame, text="Total Moves: 0")
        self.total_moves_label.pack(side='left', padx=10)
        
        # 총 crane 갯수 레이블
        self.total_cranes_label = ttk.Label(summary_frame, text="Total Cranes: 0")
        self.total_cranes_label.pack(side='left', padx=10)
        # 계산 버튼 추가
        calculate_button = ttk.Button(summary_frame, text="Calculate Gang Split",
                                    command=self.calculate_gang_split)
        calculate_button.pack(side='left', padx=315)

        # Gang Split 라인을 표시할 프레임
        self.gang_split_lines_frame = ttk.Frame(gang_split_frame)
        self.gang_split_lines_frame.pack(pady=5, fill='x')

        # 헤더 추가
        header_frame = ttk.Frame(self.gang_split_lines_frame)
        header_frame.pack(fill='x')
        ttk.Label(header_frame, text="Crane", width=5).pack(side='left', padx=10)
        ttk.Label(header_frame, text="Cntr", width=5).pack(side='left', padx=5)

        # Gang Split 라인을 저장할 리스트
        self.gang_split_lines = []

        # 크레인 라인을 한 줄에 배치
        row_frame = ttk.Frame(self.gang_split_lines_frame)
        row_frame.pack(pady=2, fill='x')
        
        for crane_id in range(1, 9):
            frame = ttk.Frame(row_frame)
            frame.pack(side='left', padx=2, expand=True)
            
            # Crane ID (읽기 전용)
            ttk.Label(frame, text=str(crane_id), width=3).pack(side='left', padx=1)
            
            # Work Time
            work_time_var = tk.StringVar(value="")
            work_time_entry = ttk.Entry(frame, textvariable=work_time_var, width=5)
            work_time_entry.pack(side='left', padx=1)
            
            line_data = {
                "frame": frame,
                "crane_id": crane_id,
                "work_time": work_time_var
            }
            
            self.gang_split_lines.append(line_data)

        # Gang Work 섹션 추가
        gang_work_frame = ttk.LabelFrame(self.operations_tab, text="Gang Work")
        gang_work_frame.pack(pady=5, fill='x', padx=10, anchor='w')
        
        # 프레임 너비 설정
        gang_work_frame.configure(width=600)
        
        # 상단 버튼과 라벨을 위한 프레임
        top_frame = ttk.Frame(gang_work_frame)
        top_frame.pack(pady=5, fill='x', padx=10)
        
        # 붙여넣기 버튼 추가
        paste_button = ttk.Button(top_frame, text="Paste Gang Work Data",
                                command=self.handle_gang_work_paste)
        paste_button.pack(side='left', padx=5)
        
        # 빈 영역 클릭 시 붙여넣기를 위한 레이블 추가
        empty_area = ttk.Label(top_frame, text="Click here to paste data")
        empty_area.pack(side='left', padx=5)
        
        # 빈 영역 클릭 이벤트 바인딩
        empty_area.bind('<Button-1>', lambda e: self.handle_gang_work_paste())
        
        # 구분선 추가
        separator = ttk.Separator(gang_work_frame, orient='horizontal')
        separator.pack(fill='x', padx=10, pady=5)
        
        # Gang Work 라인을 표시할 프레임 (스크롤 가능한 영역)
        container_frame = ttk.Frame(gang_work_frame)
        container_frame.pack(pady=5, fill='both', expand=True, padx=10)
        
        # 스크롤바 추가
        canvas = tk.Canvas(container_frame)
        scrollbar = ttk.Scrollbar(container_frame, orient="vertical", command=canvas.yview)
        self.gang_work_lines_frame = ttk.Frame(canvas)
        
        # 스크롤 가능한 영역 설정
        self.gang_work_lines_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.gang_work_lines_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 스크롤바와 캔버스 배치
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # 헤더 추가
        header_frame = ttk.Frame(self.gang_work_lines_frame)
        header_frame.pack(fill='x')
        ttk.Label(header_frame, text="Count", width=6).pack(side='left', padx=2)
        ttk.Label(header_frame, text="Start Time", width=15).pack(side='left', padx=2)
        ttk.Label(header_frame, text="End Time", width=15).pack(side='left', padx=2)

        # Gang Work 라인을 저장할 리스트
        self.gang_work_lines = []
        
        # 클립보드 바인딩 추가 (프레임 전체에 적용)
        gang_work_frame.bind('<Control-v>', lambda e: self.handle_gang_work_paste())
        empty_area.bind('<Control-v>', lambda e: self.handle_gang_work_paste())

    def setup_discharge_tab(self):
        # 상단 버튼과 라벨을 위한 프레임
        top_frame = ttk.Frame(self.discharge_tab)
        top_frame.pack(pady=5, fill='x', padx=10)
        
        # 붙여넣기 버튼 추가
        paste_button = ttk.Button(top_frame, text="Paste Container Data",
                                command=lambda: self.handle_paste('discharge'))
        paste_button.pack(side='left', padx=5)
        
        # 빈 영역 클릭 시 붙여넣기를 위한 레이블 추가
        empty_area = ttk.Label(top_frame, text="Click here to paste data")
        empty_area.pack(side='left', padx=5)
        
        # 빈 영역 클릭 이벤트 바인딩
        empty_area.bind('<Button-1>', lambda e: self.handle_paste('discharge'))
        
        # 구분선 추가
        separator = ttk.Separator(self.discharge_tab, orient='horizontal')
        separator.pack(fill='x', padx=10, pady=5)
        
        # 컨테이너 라인 프레임 (스크롤 가능한 영역)
        container_frame = ttk.Frame(self.discharge_tab)
        container_frame.pack(pady=5, fill='both', expand=True, padx=10)
        
        # 스크롤바 추가
        canvas = tk.Canvas(container_frame)
        scrollbar = ttk.Scrollbar(container_frame, orient="vertical", command=canvas.yview)
        self.discharge_lines_frame = ttk.Frame(canvas)
        
        # 스크롤 가능한 영역 설정
        self.discharge_lines_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.discharge_lines_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 스크롤바와 캔버스 배치
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # 초기 컨테이너 라인은 생성하지 않음
        self.discharge_lines = []
        
        # 클립보드 바인딩 추가 (프레임 전체에 적용)
        self.discharge_tab.bind('<Control-v>', lambda e: self.handle_paste('discharge'))
        empty_area.bind('<Control-v>', lambda e: self.handle_paste('discharge'))

    def setup_load_tab(self):
        # 상단 버튼과 라벨을 위한 프레임
        top_frame = ttk.Frame(self.load_tab)
        top_frame.pack(pady=5, fill='x', padx=10)
        
        # 붙여넣기 버튼 추가
        paste_button = ttk.Button(top_frame, text="Paste Container Data",
                                command=lambda: self.handle_paste('load'))
        paste_button.pack(side='left', padx=5)
        
        # 빈 영역 클릭 시 붙여넣기를 위한 라벨 추가
        empty_area = ttk.Label(top_frame, text="Click here to paste data")
        empty_area.pack(side='left', padx=5)
        
        # 빈 영역 클릭 이벤트 바인딩
        empty_area.bind('<Button-1>', lambda e: self.handle_paste('load'))
        
        # 구분선 추가
        separator = ttk.Separator(self.load_tab, orient='horizontal')
        separator.pack(fill='x', padx=10, pady=5)
        
        # 컨테이너 라인 프레임 (스크롤 가능한 영역)
        container_frame = ttk.Frame(self.load_tab)
        container_frame.pack(pady=5, fill='both', expand=True, padx=10)
        
        # 스크롤바 추가
        canvas = tk.Canvas(container_frame)
        scrollbar = ttk.Scrollbar(container_frame, orient="vertical", command=canvas.yview)
        self.load_lines_frame = ttk.Frame(canvas)
        
        # 스크롤 가능한 영역 설정
        self.load_lines_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.load_lines_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 스크롤바와 캔버스 배치
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # 초기 컨테이너 라인은 생성하지 않음
        self.load_lines = []
        
        # 클립보드 바인딩 추가 (프레임 전체에 적용)
        self.load_tab.bind('<Control-v>', lambda e: self.handle_paste('load'))
        empty_area.bind('<Control-v>', lambda e: self.handle_paste('load'))

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
                    values=["20", "40", "45"], state="readonly", width=4).pack(side='left', padx=2)
        
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
        
        headers = ["Account", "Type", "Size", "#", "F/E", "OOG", "RF", "IMO", "Reason" , "Not MSC A/C"]
        widths = [11, 15, 10, 5, 9, 6, 5, 5, 65, 15]
        
        for header, width in zip(headers, widths):
            ttk.Label(header_frame, text=header, width=width).pack(side='left', padx=2)

        self.shifting_lines = []
        self.add_shifting_line()  # 초기 라인 하나 추가

    def add_shifting_line(self):
        frame = ttk.Frame(self.shifting_lines_frame)
        frame.pack(pady=5, fill='x')
        
        # Get values from the line immediately above (the last line) if it exists
        previous_line_values = {}
        if self.shifting_lines:
            previous_line = self.shifting_lines[-1]  # Get the last line instead of the first line
            previous_line_values = {
                "account": previous_line["account"].get(),
                "type": previous_line["type"].get(),
                "size": previous_line["size"].get(),
                "value": previous_line["value"].get(),
                "fe": previous_line["fe"].get(),
                "oog": previous_line["oog"].get(),
                "reefer": previous_line["reefer"].get(),
                "imo": previous_line["imo"].get(),
                "reason": previous_line["reason"].get(),
                "notformscaccount": previous_line["notformscaccount"].get()
            }
        
        # Account
        account_var = tk.StringVar(value=previous_line_values.get("account", "MSCU"))
        ttk.Combobox(frame, textvariable=account_var,
                    values=["MSCU","ZIMU","HDMU","HLCU","MAEU"], state="readonly", width=8).pack(side='left', padx=2)
        
        # Type
        type_var = tk.StringVar(value=previous_line_values.get("type", "Restow"))
        # Type radio buttons frame
        type_frame = ttk.Frame(frame)
        type_frame.pack(side='left', padx=2)
        
        ttk.Radiobutton(type_frame, text="Restow", variable=type_var, value="Restow").pack(side='left')
        ttk.Radiobutton(type_frame, text="Shift", variable=type_var, value="Shift").pack(side='left')
        
        # Container Size
        size_var = tk.StringVar(value=previous_line_values.get("size", "40"))
        # Size radio buttons frame
        size_frame = ttk.Frame(frame)
        size_frame.pack(side='left', padx=2)
        ttk.Radiobutton(size_frame, text="20", variable=size_var, value="20").pack(side='left')
        ttk.Radiobutton(size_frame, text="40", variable=size_var, value="40").pack(side='left')
        
        # Value
        value_entry = ttk.Entry(frame, width=4)  # 컨테이너 수량
        value_entry.pack(side='left', padx=2)
        if "value" in previous_line_values:
            value_entry.insert(0, previous_line_values["value"])
        
        # Full/Empty
        fe_var = tk.StringVar(value=previous_line_values.get("fe", "F"))
        # F/E radio buttons frame
        fe_frame = ttk.Frame(frame)
        fe_frame.pack(side='left', padx=10)
        ttk.Radiobutton(fe_frame, text="F", variable=fe_var, value="F").pack(side='left')
        ttk.Radiobutton(fe_frame, text="E", variable=fe_var, value="E").pack(side='left')
        
        # OOG
        oog_var = tk.StringVar(value=previous_line_values.get("oog", "0"))
        oog_check = ttk.Checkbutton(frame, variable=oog_var, onvalue="1", offvalue="0")
        oog_check.pack(side='left', padx=10)
        
        # Reefer
        reefer_var = tk.StringVar(value=previous_line_values.get("reefer", "0"))
        reefer_check = ttk.Checkbutton(frame, variable=reefer_var, onvalue="1", offvalue="0")
        reefer_check.pack(side='left', padx=10)
        
        # IMO
        imo_var = tk.StringVar(value=previous_line_values.get("imo", "0")) 
        imo_check = ttk.Checkbutton(frame, variable=imo_var, onvalue="1", offvalue="0")
        imo_check.pack(side='left', padx=10)
        
        # Reason
        reason_var = tk.StringVar(value=previous_line_values.get("reason", "Restow of optional cargo onboard, to maximize vsl capacity"))
        ttk.Combobox(frame, textvariable=reason_var,
                    values=[
                        "Due To A Change Of Destination Request ( COD )",
                        "Due To A Terminal Port Crane Restriction In The Next Ports",
                        "Due To A Terminal Shore Crane Break Down",
                        "Due To A Terminal Wrong Or Bad Stowage Onboard",
                        "Due To Cargo Or Container Movement / Damage Onboard",
                        "Due To Change Of Port Rotation Or Port Omission",
                        "Due To Change Of Service And ROB Being Kept Onboard For FPOD ( T/S Cost Saving )",
                        "Due To IMO Segregation Rules Or Restrictions Onboard",
                        "Due To Late Arriving Cargo Being Stowed Whenever Posible To Avoid Delays In Operations",
                        "Due To Late Arriving Cargo Being Stowed Wherever Possible To Avoid Delays In Operations",
                        "Due To Lashing Forces Being Exceeded",
                        "Due To Local Port Restrictions Or Customs Inspections To Avoid Fines At POD",
                        "Due To Mis-Routing Declaration Of Cargo",
                        "Due To Reefer Segregation Onboard ( IMO - Reefer Conflicts )",
                        "Due To Sailing Deadlines / Port Strikes Imposed Or Cut And Runs At The Last Minute",
                        "Due To The Vessel Or Vessels Equipment Being Damaged Or Failing",
                        "Empty Unit Repositioning",
                        "FNTE Planning Oversight",
                        "IMO Constraint - Due To IMO Segregation Rules Or Restrictions Onboard",
                        "Marine Operations Restows Request",
                        "Port Restrictions Due To Reefer Handling Capacity Ashore",
                        "Reefer Malfunction While On Board",
                        "Reefer Restows Performed To Maximise Vessel Reefer Capacity",
                        "Reefer Units Loaded In Temporary Overstow Positions To Allow Connection To Reefer Sockets",
                        "Re-Stow For Operational Reason. Shared Cost As Per 2M Contract (% Of Throughput)",
                        "Restow Carried Out For Ships Spares / Power Pack",
                        "Restow Of Optional Cargo Onboard, To Maximize Vsl Capacity",
                        "Restows Performed To Assist In Maximising Vessel Capacity",
                        "Special Cargo Overstowing Import Cargo To Be Discharged (B.Bulk / OOG / Sensitive Commodities)",
                        "Terminal Convenience",
                        "To Create Additional Space For Loading Heavy Containers",
                        "To Keep The Vessels Stability /Stress / Visibility Parameters Within Safety Limits",
                        "To Reduce The Deck Tier Unit Height To Avoid Extra Transit Surcharges Crossing The Canals",
                        "Vessel Diverted Out Of Suez To Avoid Red Sea Risk Area"
                    ], 
                    state="readonly", width=60).pack(side='left', padx=10)
        
        # Not For MSC Account
        notformscaccount_var = tk.BooleanVar(value=previous_line_values.get("notformscaccount", False))
        notformscaccount_check = ttk.Checkbutton(frame, variable=notformscaccount_var)
        notformscaccount_check.pack(side='left', padx=40)
        
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
        """XML 생성 함수"""
        try:
            # Gang Split 계산 여부 확인
            if not self.gang_split_calculated:
                messagebox.showerror("Error", "Gang Split을 먼저 계산해주세요.")
                return
                
            # XML 생성기 초기화
            generator = VCIXMLGenerator()
            
            # Header 정보 추가
            generator.add_header(
                vessel=self.vessel_entry.get(),
                voyage=self.voyage_entry.get(),
                portun=self.portun_entry.get(),
                master=".",  # 고정값
                berth="1",   # 고정값
                general_remark=self.general_remark_text.get("1.0", "end-1c")
            )
            
            # Arrival 정보 추가
            arrival = SubElement(generator.root, 'arrival')
            
            # Timeline 섹션
            timeline = SubElement(arrival, 'timeline')
            self.add_timeline_field(timeline, 'BOWTHARR', 'S', self.bowtharr_var.get())
            self.add_timeline_field(timeline, 'ARRPILSTA', 'D', 
                self.convert_datetime(self.arrpil_time_entry.get()))
            self.add_timeline_field(timeline, 'REAFORANC', 'S', self.arr_reaforanc_var.get())
            
            # Draft 섹션
            draft = SubElement(arrival, 'draft')
            self.add_draft_item(draft, 'AFT', self.arr_draft_aft_entry.get())
            self.add_draft_item(draft, 'FWD', self.arr_draft_fwd_entry.get())
            
            # Pilots 섹션
            pilots = SubElement(arrival, 'pilots')
            pilots.set('cancelled', 'false')
            pilot = SubElement(pilots, 'pilot')
            pilot.set('type', 'Sea')
            pilot.set('number', '1')
            pilot.set('from', self.convert_datetime(self.arr_pilot_from_entry.get()))
            pilot.set('to', self.convert_datetime(self.arr_pilot_to_entry.get()))
            
            # Towages 섹션
            towages = SubElement(arrival, 'towages')
            towages.set('cancelled', 'false')
            
            for i, entries in enumerate(self.arr_tug_entries):
                tug = SubElement(towages, 'tug')
                tug.set('type', 'Sea')
                tug.set('number', str(i + 1))
                tug.set('from', self.convert_datetime(entries['From'].get()))
                tug.set('to', self.convert_datetime(entries['To'].get()))
                tug.set('comment', entries['Comment'].get())
                tug.set('tugtype', 'Conventional')
                tug.set('name', ' ')
                tug.set('bowthrusternonop', '0')
            
            # Operations 정보 추가
            operations = SubElement(generator.root, 'operations')
            
            # Timeline 섹션
            timeline = SubElement(operations, 'timeline')
            self.add_timeline_field(timeline, 'DOCSIDTO', 'S', self.docsidto_var.get())
            self.add_timeline_field(timeline, 'DOCATTER', 'D', 
                self.convert_datetime(self.docatter_entry.get()))
            self.add_timeline_field(timeline, 'GANWAYDOWN', 'D', 
                self.convert_datetime(self.ganwaydown_entry.get()))
            self.add_timeline_field(timeline, 'OPECOM', 'D', 
                self.convert_datetime(self.opecom_entry.get()))
            self.add_timeline_field(timeline, 'OPECOMP', 'D', 
                self.convert_datetime(self.opecomp_entry.get()))
            self.add_timeline_field(timeline, 'LASCOMP', 'D', 
                self.convert_datetime(self.lascomp_entry.get()))
            self.add_timeline_field(timeline, 'LASCOMPBY', 'S', self.lascompby_var.get())
            
            # Departure 정보 추가
            departure = SubElement(generator.root, 'departure')
            
            # Timeline 섹션
            timeline = SubElement(departure, 'timeline')
            self.add_timeline_field(timeline, 'BOWTHDEP', 'S', self.bowthdep_var.get())
            self.add_timeline_field(timeline, 'VESUNDOC', 'D', 
                self.convert_datetime(self.vesundoc_entry.get()))
            self.add_timeline_field(timeline, 'REAFORANC', 'S', self.dep_reaforanc_var.get())
            
            # Draft 섹션
            draft = SubElement(departure, 'draft')
            self.add_draft_item(draft, 'AFT', self.dep_draft_aft_entry.get())
            self.add_draft_item(draft, 'FWD', self.dep_draft_fwd_entry.get())
            
            # Pilots 섹션
            pilots = SubElement(departure, 'pilots')
            pilots.set('cancelled', 'false')
            pilot = SubElement(pilots, 'pilot')
            pilot.set('type', 'Sea')
            pilot.set('number', '1')
            pilot.set('from', self.convert_datetime(self.dep_pilot_from_entry.get()))
            pilot.set('to', self.convert_datetime(self.dep_pilot_to_entry.get()))
            
            # Towages 섹션
            towages = SubElement(departure, 'towages')
            towages.set('cancelled', 'false')
            
            for i, entries in enumerate(self.dep_tug_entries):
                tug = SubElement(towages, 'tug')
                tug.set('type', 'Sea')
                tug.set('number', str(i + 1))
                tug.set('from', self.convert_datetime(entries['From'].get()))
                tug.set('to', self.convert_datetime(entries['To'].get()))
                tug.set('comment', entries['Comment'].get())
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
                linecode.set('terminal', self.portun_entry.get())
            
            # Load Details
            load = SubElement(generator.root, 'loaddetails')
            for line in self.load_lines:
                linecode = SubElement(load, 'linecode')
                linecode.set('type', line['type'].get())
                linecode.set('operator', line['operator'].get())
                linecode.set('containersize', line['size'].get())
                linecode.set('fullempty', line['fe'].get())
                linecode.set('number', line['number'].get())
                linecode.set('terminal', self.portun_entry.get())
            
            # Lid Moves
            lidmoves = SubElement(generator.root, 'lidmoves')
            lid = SubElement(lidmoves, 'lid')
            lid.set('terminal', self.portun_entry.get())
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
                shift.set('oog', 'true' if line['oog'].get() == "1" else 'false')
                shift.set('reefer', 'true' if line['reefer'].get() == "1" else 'false')
                shift.set('imo', 'true' if line['imo'].get() == "1" else 'false')
                shift.set('notformscaccount', 'true' if line['notformscaccount'].get() else 'false')
            
            # Gang Split 추가
            gangsplit = SubElement(generator.root, 'gangsplit')
            for line in self.gang_split_lines:
                crane = SubElement(gangsplit, 'crane')
                crane.set('craneid', str(line['crane_id']))
                crane.set('final', line['work_time'].get())
            
            # Gang Work Time 추가
            gangworktime = SubElement(generator.root, 'gangworkingtimes')
            for line in self.gang_work_lines:
                gang = SubElement(gangworktime, 'gang')
                gang.set('number', line['count'].get())
                
                # datetime 형식으로 변환
                from_time = datetime.strptime(line['start_time'].get(), "%Y%m%d %H%M")
                to_time = datetime.strptime(line['end_time'].get(), "%Y%m%d %H%M")
                
                gang.set('from', from_time.strftime("%Y-%m-%dT%H:%M:00"))
                gang.set('to', to_time.strftime("%Y-%m-%dT%H:%M:00"))
                gang.set('workingperiodpayrate', "ST")
                gang.set('terminal', "KRPUSPN")

            # Summary
            summary = SubElement(generator.root, 'summary')
            
            # Gang Summary
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
            for name, entry in self.terminal_efficiency_entries.items():
                item = SubElement(terminal_efficiency, name)
                value = entry.get()
                
                # 날짜/시간 필드에 대한 특별 처리
                if name in ['Planner_Time_Msc', 'Planner_Time_LP', 'Terminal_Time'] and value:
                    try:
                        value = self.convert_datetime(value)
                    except Exception as e:
                        print(f"Error converting datetime for {name}: {e}")
                
                item.text = value or "0"  # 값이 없으면 "0" 사용
            
            # XML 파일 생성
            tree = ElementTree(generator.root)
            filename = f"{self.vessel_entry.get()}_{self.voyage_entry.get()}_{self.portun_entry.get()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xml"
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
                
            results = parse_excel_data(clipboard_data)
            if not results:
                messagebox.showwarning("Warning", "파싱할 수 있는 컨테이너 데이터가 없습니다.")
                return
            
            # 새 데이터로 컨테이너 라인 생성
            container_data = []
            new_operator = clipboard_data.strip().split('\n')[0].split('\t')[0]
            
            # DIMP 데이터 처리
            for fe in ['FULL', 'EMPTY']:
                for size in ['20', '40', '45']:
                    count = results['DIMP'][fe][size]
                    if count > 0:
                        container_data.append({
                            'type': 'DIMP' if tab_type == 'discharge' else 'DEXP',
                            'operator': new_operator,
                            'containersize': int(size),
                            'fullempty': 'F' if fe == 'FULL' else 'E',
                            'number': count,
                            'terminal': 'PTSIEPS'
                        })
            
            # TRAN 데이터 처리
            for fe in ['FULL', 'EMPTY']:
                for size in ['20', '40', '45']:
                    count = results['TRAN'][fe][size]
                    if count > 0:
                        container_data.append({
                            'type': 'TRAN',
                            'operator': new_operator,
                            'containersize': int(size),
                            'fullempty': 'F' if fe == 'FULL' else 'E',
                            'number': count,
                            'terminal': 'PTSIEPS'
                        })
            
            # 컨테이너 라인 생성 (기존 데이터는 유지하고 새 데이터만 추가)
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
            
            # 시간 순으로 정렬하고 중복 제거
            all_times = sorted(list(set(all_times)))

            # 각 시간 구간별 작동 중인 QC 수 계산
            time_periods = []
            for i in range(len(all_times) - 1):
                current_time = all_times[i]
                next_time = all_times[i + 1]
                
                # 현재 시간대에 작동 중인 QC 수 계산
                active_qcs = sum(1 for start, end in qc_times 
                               if start <= current_time and end > current_time)
                
                if active_qcs > 0:  # QC가 작동 중인 경우만 포함
                    # 시작 시간에 1분 추가
                    if i > 0:
                        current_time = current_time + timedelta(minutes=1)
                    time_periods.append((active_qcs, current_time, next_time))

            # 트리뷰 초기화
            for item in self.qc_tree.get_children():
                self.qc_tree.delete(item)

            # 결과 데이터 트리뷰에 추가
            for count, start, end in time_periods:
                self.qc_tree.insert('', 'end', values=(
                    str(count),
                    start.strftime("%Y%m%d %H%M"),
                    end.strftime("%Y%m%d %H%M")
                ))

            # 디버깅을 위한 파싱된 데이터 출력
            print("\n생성된 결과:")
            for count, start, end in time_periods:
                print(f"QC 수: {count}, 시작: {start.strftime('%Y-%m-%d %H:%M')}, 종료: {end.strftime('%Y-%m-%d %H:%M')}")

            messagebox.showinfo("Success", f"{len(time_periods)}개의 시간대별 QC 작업 데이터가 처리되었습니다.")

        except Exception as e:
            error_msg = f"QC 데이터 처리 중 오류 발생: {str(e)}"
            print(f"Error details: {str(e)}")
            messagebox.showerror("Error", error_msg)

    def setup_terminal_efficiency_tab(self):
        """
        Terminal Efficiency 탭 설정
        """
        # 스크롤 가능한 프레임 생성
        main_frame = ttk.Frame(self.terminal_efficiency_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # 스크롤바 추가
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 스크롤바와 캔버스 배치
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # Terminal Efficiency 항목들
        efficiency_items = [
            ('Planner_Time_Msc', 'CLL'),
            ('Planner_Time_LP', 'OBL'),
            ('Terminal_Time', 'Plan to Terminal time'),
            ('CI_Proforma', 'Proforma Cranes'),
            ('CI_Planner', 'Best possible from the planner'),
            ('CI_Gang_Availability', 'Provided by the planner '),
            ('CI_Sub_Optimal', 'Reason for being sub-optimal'),
            ('Starting_Operations', 'Were changes made to the load list less that  12 hours before starting operations?'),
            ('Quantity_Of_Containers', 'How many containers rolled from this vessel to another within 12 hours of operations start?'),
            ('Changes_After_Ingate', 'How many containers loaded on this vessel had their Vessel, POD, Weight changes after ingate?'),
            ('Moved_from_Another_Terminal', 'How many containers loaded on this vessel had to be moved from another terminal in this port?'),
            ('Qty_of_Containers_before_Arrival', 'How many containers were not on the quay 6 hours before arrival?'),
            ('Live_Connections', 'Number of live connections to the export vessel'),
            ('Reason_for_Live_Connections', 'Reason for Live Connections'),
            ('Affect_by_Cut_And_Run', 'Number of containers affected by Cut and Run'),
            ('Reason_For_Cut_and_Run', 'Reason For Cut and Run'),
            ('Average_Yard_Utilisation', 'Average Yard Utilisation'),
            ('Remarks', 'Remarks')
        ]
        
        # 각 항목에 대한 입력 필드 생성
        self.terminal_efficiency_entries = {}
        
        for item_id, label_text in efficiency_items:
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill='x', pady=5)
            
            # 레이블
            ttk.Label(frame, text=label_text, width=100).pack(side='right')
            
            # 입력 필드 - datetime 필드는 PlaceholderEntry 사용
            if item_id in ['Planner_Time_Msc', 'Planner_Time_LP']:
                entry = PlaceholderEntry(frame, placeholder="YYYYMMDD HHMM", width=40)
                entry.bind('<KeyRelease>', lambda e: self.handle_date_change(entry))
                self.all_datetime_entries.append(entry)
            else:
                entry = ttk.Entry(frame, width=40)
            entry.pack(side='left', padx=5)
            
            # 딕셔너리에 저장
            self.terminal_efficiency_entries[item_id] = entry
        
        # 붙여넣기 버튼 추가
        paste_button = ttk.Button(scrollable_frame, text="Paste Terminal Efficiency Data",
                                command=self.handle_terminal_efficiency_paste)
        paste_button.pack(pady=10)
        
        # 빈 영역 클릭 시 붙여넣기를 위한 레이블 추가
        empty_area = ttk.Label(scrollable_frame, text="Click here to paste data")
        empty_area.pack(pady=10, expand=True, fill='both')
        
        # 빈 영역 클릭 이벤트 바인딩
        empty_area.bind('<Button-1>', lambda e: self.handle_terminal_efficiency_paste())

    def handle_terminal_efficiency_paste(self):
        """
        Terminal Efficiency 탭에 클립보드 데이터 붙여넣기 처리
        """
        try:
            clipboard_data = self.root.clipboard_get()
            if not clipboard_data.strip():
                messagebox.showwarning("Warning", "클립보드가 비어있습니다.")
                return

            # 데이터 매핑 딕셔너리 정의
            field_mapping = {
                "Plan to Terminal time": "Terminal_Time",
                "Proforma Cranes": "CI_Proforma",
                "Best possible from the planner": "CI_Planner",
                "Provided by the planner in line with gang availability": "CI_Gang_Availability",
                "Reason for being sub-optimal": "CI_Sub_Optimal",
                "Were changes made to the load list less that  12 hours before starting operations?": "Starting_Operations",
                "How many containers rolled from this vessel to another within 12 hours of operations start?": "Quantity_Of_Containers",
                "How many containers loaded on this vessel had their Vessel, POD, Weight changes after ingate?": "Changes_After_Ingate",
                "How many containers loaded on this vessel had to be moved from another terminal in this port?": "Moved_from_Another_Terminal",
                "How many containers were not on the quay 6 hours before arrival?": "Qty_of_Containers_before_Arrival",
                "Number of live connections to the export vessel": "Live_Connections",
                "Reason for live connections": "Reason_for_Live_Connections",
                "Number of containers affected by Cut and Run": "Affect_by_Cut_And_Run",
                "Reason for cut and run": "Reason_For_Cut_and_Run",
                "Average yard utilisation during port stay": "Average_Yard_Utilisation",
                "Port Stay remarks": "Remarks"
            }

            # 0으로 처리해야 하는 필드들
            zero_fields = ["Quantity_Of_Containers", "Changes_After_Ingate", 
                         "Moved_from_Another_Terminal", "Qty_of_Containers_before_Arrival",
                         "Live_Connections", "Affect_by_Cut_And_Run"]

            # 클립보드 데이터 파싱
            lines = clipboard_data.strip().split('\n')
            for line in lines:
                parts = [part.strip() for part in line.split('\t') if part.strip()]
                if len(parts) >= 2:
                    field_name = parts[0].strip()
                    field_value = parts[1].strip()

                    # 매핑된 필드 찾기
                    mapped_field = None
                    for key, value in field_mapping.items():
                        if key.lower() in field_name.lower():
                            mapped_field = value
                            break

                    if mapped_field and mapped_field in self.terminal_efficiency_entries:
                        entry = self.terminal_efficiency_entries[mapped_field]
                        entry.delete(0, tk.END)

                        # 날짜/시간 필드 특별 처리
                        if mapped_field in ["Planner_Time_Msc", "Planner_Time_LP", "Terminal_Time"]:
                            try:
                                dt = datetime.strptime(field_value, "%Y-%m-%d %H:%M")
                                formatted_value = dt.strftime("%Y%m%d %H%M")
                                entry.insert(0, formatted_value)
                            except ValueError:
                                entry.insert(0, field_value)
                        # CI_Proforma, CI_Planner, CI_Gang_Availability 필드에 대한 반올림 처리
                        elif mapped_field in ["CI_Proforma", "CI_Planner", "CI_Gang_Availability"]:
                            try:
                                value = float(field_value)
                                rounded_value = round(value)
                                entry.insert(0, str(rounded_value))
                            except ValueError:
                                entry.insert(0, field_value)
                        # CI_Sub_Optimal 필드에 대한 특별 처리
                        elif mapped_field == "CI_Sub_Optimal":
                            try:
                                # 문자열을 정수로 변환
                                value = int(field_value)
                                # 숫자에 따른 텍스트 매핑
                                sub_optimal_mapping = {
                                    1: "Crane intensity was optimal",
                                    2: "Insufficient gangs",
                                    3: "Other vessels prioritized",
                                    4: "Saving overtime costs",
                                    5: "Vessel had time"
                                }
                                # 매핑된 텍스트를 입력 필드에 표시
                                if value in sub_optimal_mapping:
                                    entry.insert(0, sub_optimal_mapping[value])
                                else:
                                    entry.insert(0, field_value)
                            except ValueError:
                                entry.insert(0, field_value)
                        # Average Yard Utilisation 필드에 대한 특별 처리
                        elif mapped_field == "Average_Yard_Utilisation":
                            try:
                                # 숫자와 소수점만 추출
                                numeric_value = ''.join(c for c in field_value if c.isdigit() or c == '.')
                                if numeric_value:
                                    # 소수점이 있는 경우 소수점 첫째자리까지만 표시
                                    if '.' in numeric_value:
                                        numeric_value = str(round(float(numeric_value), 1))
                                    entry.insert(0, numeric_value)
                                else:
                                    entry.insert(0, "0")
                            except ValueError:
                                entry.insert(0, "0")
                        # Reason_for_Live_Connections 필드에 대한 특별 처리
                        elif mapped_field == "Reason_for_Live_Connections":
                            try:
                                # 문자열을 정수로 변환
                                value = int(field_value)
                                # 숫자에 따른 텍스트 매핑
                                live_connections_mapping = {
                                    1: "No live connections",
                                    2: "Commercial request",
                                    3: "Network need",
                                    4: "Inbound vessel delay"
                                }
                                # 매핑된 텍스트를 입력 필드에 표시
                                if value in live_connections_mapping:
                                    entry.insert(0, live_connections_mapping[value])
                                else:
                                    entry.insert(0, field_value)
                            except ValueError:
                                entry.insert(0, field_value)
                        # Reason_For_Cut_and_Run 필드에 대한 특별 처리
                        elif mapped_field == "Reason_For_Cut_and_Run":
                            try:
                                # 문자열을 정수로 변환
                                value = int(field_value)
                                # 숫자에 따른 텍스트 매핑
                                cut_and_run_mapping = {
                                    1: "No Cut and run",
                                    2: "bad weather",
                                    3: "Berth competition",
                                    4: "Canal Passage",
                                    5: "Cargo Readiness",
                                    6: "Crane Breakdown",
                                    7: "Lack of gangs",
                                    8: "Port close",
                                    9: "Productivity issues",
                                    10: "Safety and Security",
                                    11: "Schedule integrity",
                                    12: "Terminal congestion",
                                    13: "Tide"
                                }
                                # 매핑된 텍스트를 입력 필드에 표시
                                if value in cut_and_run_mapping:
                                    entry.insert(0, cut_and_run_mapping[value])
                                else:
                                    entry.insert(0, field_value)
                            except ValueError:
                                entry.insert(0, field_value)
                        # Starting_Operations 필드에 대한 특별 처리
                        elif mapped_field == "Starting_Operations":
                            try:
                                value = int(field_value)
                                starting_operations_mapping = {
                                    1: "No changes made",
                                    2: "Additions",
                                    3: "Cuts",
                                    4: "Changes",
                                    5: "Additions, Cuts and or Changes"
                                }
                                if value in starting_operations_mapping:
                                    entry.insert(0, starting_operations_mapping[value])
                                else:
                                    entry.insert(0, field_value)
                            except ValueError:
                                entry.insert(0, field_value)
                        # 0으로 처리해야 하는 필드들 처리
                        elif mapped_field in zero_fields:
                            if not field_value or field_value.strip().upper() in ["NIL", "NIL", "NIL", "0"]:
                                entry.insert(0, "0")
                            else:
                                entry.insert(0, field_value)
                        else:
                            entry.insert(0, field_value)

            messagebox.showinfo("Success", "Terminal Efficiency 데이터가 성공적으로 입력되었습니다.")
            
        except Exception as e:
            error_msg = f"Terminal Efficiency 데이터 처리 중 오류 발생: {str(e)}"
            print(f"Error details: {str(e)}")  # 디버깅을 위한 상세 에러 출력
            messagebox.showerror("Error", error_msg)

    def update_departure_times(self, event=None):
        """Sailed From Berth 값이 변경될 때마다 Departure 탭의 Pilot, Tug 값을 업데이트"""
        try:
            # 입력된 값 가져오기
            sailed_str = self.sailed_from_berth_entry.get()
            if not sailed_str:
                return
            
            print(f"Processing sailed_str: {sailed_str}")
            
            # 입력 형식에 따라 파싱 (YYYYMMDD HHMM 또는 YYYY-MM-DDThh:mm:ss)
            if len(sailed_str) == 13 and ' ' in sailed_str:  # YYYYMMDD HHMM 형식
                date_part = sailed_str[:8]
                time_part = sailed_str[9:]
                year = date_part[:4]
                month = date_part[4:6]
                day = date_part[6:8]
                hour = time_part[:2]
                minute = time_part[2:4]
                sailed_time = datetime(int(year), int(month), int(day), int(hour), int(minute))
            else:  # ISO 형식
                sailed_time = datetime.strptime(sailed_str, "%Y-%m-%dT%H:%M:%S")
            
            print(f"Parsed sailed_time: {sailed_time}")
            
            # 1. Pilot 시간 계산
            pilot_from = sailed_time - timedelta(hours=1)  # Sailed From Berth - 1시간
            pilot_to = sailed_time + timedelta(minutes=30)  # Sailed From Berth + 30분
            
            # Pilot 시간 설정
            if hasattr(self, 'dep_pilot_from_entry') and self.dep_pilot_from_entry is not None:
                self.dep_pilot_from_entry.delete(0, tk.END)
                self.dep_pilot_from_entry.insert(0, pilot_from.strftime("%Y%m%d %H%M"))
                print(f"Updated Departure Pilot From to: {pilot_from.strftime('%Y%m%d %H%M')}")
            else:
                print("dep_pilot_from_entry is not initialized")
            
            if hasattr(self, 'dep_pilot_to_entry') and self.dep_pilot_to_entry is not None:
                self.dep_pilot_to_entry.delete(0, tk.END)
                self.dep_pilot_to_entry.insert(0, pilot_to.strftime("%Y%m%d %H%M"))
                print(f"Updated Departure Pilot To to: {pilot_to.strftime('%Y%m%d %H%M')}")
            else:
                print("dep_pilot_to_entry is not initialized")
            
            # 2. Tug 시간 계산
            tug_from = sailed_time - timedelta(hours=1)  # Sailed From Berth - 1시간
            tug_to = sailed_time + timedelta(hours=1)  # Sailed From Berth + 1시간
            
            # Tug 시간 설정
            if hasattr(self, 'dep_tug_entries') and self.dep_tug_entries:
                for tug_entry in self.dep_tug_entries:
                    tug_entry["From"].delete(0, tk.END)
                    tug_entry["From"].insert(0, tug_from.strftime("%Y%m%d %H%M"))
                    
                    tug_entry["To"].delete(0, tk.END)
                    tug_entry["To"].insert(0, tug_to.strftime("%Y%m%d %H%M"))
                
                print(f"Updated Departure Tug From to: {tug_from.strftime('%Y%m%d %H%M')}")
                print(f"Updated Departure Tug To to: {tug_to.strftime('%Y%m%d %H%M')}")
            else:
                print("dep_tug_entries is not initialized or empty")
            
        except Exception as e:
            print(f"Error updating departure times: {str(e)}")
            import traceback
            traceback.print_exc()

def main():
    root = tk.Tk()
    app = VCIGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
