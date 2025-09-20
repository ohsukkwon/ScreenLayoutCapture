# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import subprocess
import os
import sys
from datetime import datetime
from PIL import Image, ImageTk, ImageDraw
import xml.etree.ElementTree as ET
import xml.dom.minidom
from io import BytesIO
import pystray
from pystray import MenuItem as item
import threading
import time
import re
from tkinter import font as tkfont
import requests
import json

is_stp_mode = False#True

# for stp_mode
if is_stp_mode:
    var_content_desc = 'talkback'
else:
    var_content_desc = 'content-desc'

g_attributes = ['text=', 'resource-id=', f'{var_content_desc}=', 'package=', 'class=', 'bounds=']


class ScreenLayoutCapture:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Screen & Layout Capture [ver.1.250914_101]")
        self.root.geometry("1200x800")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Alt+F4 키 바인딩 추가
        self.root.bind('<Alt-F4>', self.on_alt_f4)

        # 프로그램 아이콘 설정
        icon_path = os.path.join(os.path.dirname(__file__), "icon_ScreenLayoutCapture.png")
        if os.path.exists(icon_path):
            icon_image = ImageTk.PhotoImage(file=icon_path)
            self.root.iconphoto(True, icon_image)

        # 변수 초기화
        self.current_device = None
        self.current_id = ""
        self.screen_image = None
        self.layout_xml = ""
        self.selection_start = None
        self.selection_end = None
        self.selection_rect = None
        self.canvas_image = None
        self.original_image = None

        # 시스템 트레이 관련
        self.tray_icon = None
        self.is_minimized_to_tray = False

        # Layout capture 관련 변수들
        self.font_size = 10
        self.undo_stack = []
        self.redo_stack = []
        self.is_recording_change = False

        # 검색 관련 변수들
        self.search_text = ""
        self.search_positions = []
        self.current_search_index = -1
        self.last_search_end = None

        # 디바이스 탭 관련 변수
        self.device_tabs = {}  # device_id -> tab_frame 매핑

        self.setup_ui()
        self.load_devices()

        # Ctrl+X 조합 감지를 위한 상태 추적
        self.ctrl_pressed = False
        self.root.bind('<KeyPress-Control_L>', self.on_ctrl_press)
        self.root.bind('<KeyPress-Control_R>', self.on_ctrl_press)
        self.root.bind('<KeyRelease-Control_L>', self.on_ctrl_release)
        self.root.bind('<KeyRelease-Control_R>', self.on_ctrl_release)
        self.root.focus_set()

    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 상단 프레임 (버튼들) - 시스템트레이 버튼 제거
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # 탭 컨트롤 생성
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 디바이스 관리 탭
        self.setup_device_tab()

    def setup_device_tab(self):
        """디바이스 관리 탭 설정"""
        device_tab = ttk.Frame(self.notebook)
        self.notebook.add(device_tab, text="Device Management")

        # 디바이스 관리 프레임
        device_frame = ttk.LabelFrame(device_tab, text="Connected Devices")
        device_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 디바이스 목록과 새로고침 버튼
        devices_control_frame = ttk.Frame(device_frame)
        devices_control_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(devices_control_frame, text="Devices :").pack(side=tk.LEFT)
        ttk.Button(devices_control_frame, text="새로고침", command=self.load_devices).pack(side=tk.RIGHT)

        # 디바이스 리스트박스
        listbox_frame = ttk.Frame(device_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.device_listbox = tk.Listbox(listbox_frame, height=15)
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.device_listbox.yview)
        self.device_listbox.configure(yscrollcommand=scrollbar.set)

        self.device_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        #self.device_listbox.bind('<<ListboxSelect>>', self.on_device_select)
        self.device_listbox.bind('<Double-Button-1>', self.on_device_double_click)

        # 선택된 디바이스 정보
        info_frame = ttk.LabelFrame(device_frame, text="Selected Device Info")
        info_frame.pack(fill=tk.X, padx=5, pady=(10, 5))

        self.device_info_text = scrolledtext.ScrolledText(info_frame, height=200, wrap=tk.WORD)
        self.device_info_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def create_capture_tab(self, device_id):
        """특정 디바이스를 위한 화면 캡처 탭 생성"""
        capture_tab = ttk.Frame(self.notebook)

        # 탭 제목에 X 버튼 추가
        tab_title = f"Screen & Layout Capture({device_id})"
        self.notebook.add(capture_tab, text=tab_title)

        # X 버튼 이벤트 처리를 위한 바인딩
        self.notebook.bind('<Button-3>', lambda e: self.on_tab_right_click(e, capture_tab))

        # ID 프레임
        id_frame = ttk.Frame(capture_tab)
        id_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        ttk.Label(id_frame, text="ID :").pack(side=tk.LEFT)
        id_entry = ttk.Entry(id_frame, width=30)
        id_entry.pack(side=tk.LEFT, padx=(5, 0))

        # 캡처 컨트롤 프레임
        control_frame = ttk.Frame(capture_tab)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(control_frame, text="reload cap",
                   command=lambda: self.reload_capture(device_id)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="save cap",
                   command=lambda: self.save_capture(device_id)).pack(side=tk.LEFT)

        # 현재 선택된 디바이스 표시
        current_device_label = ttk.Label(control_frame, text=f"디바이스: {device_id}", foreground="green")
        current_device_label.pack(side=tk.RIGHT)

        # 메인 컨텐츠 프레임 (스크린 캡처와 레이아웃)
        content_frame = ttk.Frame(capture_tab)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # PanedWindow를 사용하여 크기 조절 가능하게 만듦
        paned_window = ttk.PanedWindow(content_frame, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)

        # 스크린 캡처 프레임
        screen_frame = ttk.LabelFrame(paned_window, text="Screen capture")
        paned_window.add(screen_frame, weight=1)

        # 축소 비율 선택 라디오 버튼 (요청사항 1)
        scale_frame = ttk.Frame(screen_frame)
        scale_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(scale_frame, text="축소 비율:").pack(side=tk.LEFT)

        scale_var = tk.StringVar(value="25")
        scale_options = [("10%", "10"), ("20%", "20"), ("25%", "25"), ("50%", "50"), ("100%", "100")]

        for text, value in scale_options:
            ttk.Radiobutton(scale_frame, text=text, variable=scale_var, value=value,
                            command=lambda did=device_id: self.on_scale_change(did)).pack(side=tk.LEFT, padx=5)

        # 좌표 표시 라벨 (요청사항 1)
        coords_label = ttk.Label(screen_frame, text="이미지 Resolution : -, 좌표 : (-,-)")
        coords_label.pack(anchor="w", padx=5, pady=(3, 0))

        # 요청사항: x,y,width,height 입력창과 표시/삭제 버튼 추가
        rect_frame = ttk.Frame(screen_frame)
        rect_frame.pack(fill=tk.X, padx=5, pady=5)

        # x,y,width,height 입력창
        ttk.Label(rect_frame, text="x:").grid(row=0, column=0, padx=(0, 5))
        x_entry = ttk.Entry(rect_frame, width=8)
        x_entry.grid(row=0, column=1, padx=(0, 10))

        ttk.Label(rect_frame, text="y:").grid(row=0, column=2, padx=(0, 5))
        y_entry = ttk.Entry(rect_frame, width=8)
        y_entry.grid(row=0, column=3, padx=(0, 10))

        ttk.Label(rect_frame, text="width:").grid(row=0, column=4, padx=(0, 5))
        width_entry = ttk.Entry(rect_frame, width=8)
        width_entry.grid(row=0, column=5, padx=(0, 10))

        ttk.Label(rect_frame, text="height:").grid(row=0, column=6, padx=(0, 5))
        height_entry = ttk.Entry(rect_frame, width=8)
        height_entry.grid(row=0, column=7, padx=(0, 10))

        # 표시/삭제 버튼
        show_rect_button = ttk.Button(rect_frame, text="표시",
                                      command=lambda: self.show_yellow_rectangle(device_id))
        show_rect_button.grid(row=0, column=8, padx=5)

        delete_rect_button = ttk.Button(rect_frame, text="삭제",
                                        command=lambda: self.delete_yellow_rectangle(device_id))
        delete_rect_button.grid(row=0, column=9, padx=5)

        # Screen capture 파일 경로 표시 라벨 추가 (요청사항)
        screen_path_label = ttk.Label(screen_frame, text="", foreground="blue")
        screen_path_label.pack(anchor="w", padx=5, pady=(3, 0))

        # 스크린 캡처 캔버스 (스크롤바 추가)
        screen_canvas_frame = ttk.Frame(screen_frame)
        screen_canvas_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        screen_canvas = tk.Canvas(screen_canvas_frame, bg='white', width=400, height=600)
        screen_v_scrollbar = ttk.Scrollbar(screen_canvas_frame, orient=tk.VERTICAL, command=screen_canvas.yview)
        screen_h_scrollbar = ttk.Scrollbar(screen_canvas_frame, orient=tk.HORIZONTAL, command=screen_canvas.xview)
        screen_canvas.configure(yscrollcommand=screen_v_scrollbar.set, xscrollcommand=screen_h_scrollbar.set)

        screen_canvas.grid(row=0, column=0, sticky="nsew")
        screen_v_scrollbar.grid(row=0, column=1, sticky="ns")
        screen_h_scrollbar.grid(row=1, column=0, sticky="ew")

        screen_canvas_frame.grid_rowconfigure(0, weight=1)
        screen_canvas_frame.grid_columnconfigure(0, weight=1)

        # 레이아웃 캡처 프레임
        layout_frame = ttk.LabelFrame(paned_window, text="Layout capture")
        paned_window.add(layout_frame, weight=1)

        # Layout capture 파일 경로 표시 라벨 추가 (요청사항)
        layout_path_label = ttk.Label(layout_frame, text="", foreground="blue")
        layout_path_label.pack(anchor="w", padx=5, pady=(3, 0))

        # 레이아웃 캡처 컨트롤 프레임 (폰트 크기, 검색)
        layout_control_frame = ttk.Frame(layout_frame)
        layout_control_frame.pack(fill=tk.X, padx=5, pady=5)

        # 폰트 크기 표시
        font_label = ttk.Label(layout_control_frame, text=f"Font : 10")
        font_label.pack(side=tk.LEFT)

        # 검색 프레임
        search_frame = ttk.Frame(layout_control_frame)
        search_frame.pack(side=tk.RIGHT)

        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT)
        search_entry = ttk.Entry(search_frame, width=20)
        search_entry.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(search_frame, text="검색",
                   command=lambda: self.search_text_in_layout(device_id)).pack(side=tk.LEFT, padx=(5, 0))

        # 레이아웃 텍스트와 검색 컴포넌트를 분리하는 PanedWindow 추가 (세로 방향)
        layout_paned_window = ttk.PanedWindow(layout_frame, orient=tk.VERTICAL)
        layout_paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 레이아웃 텍스트 프레임 (PanedWindow의 첫 번째 패널)
        layout_text_panel = ttk.Frame(layout_paned_window)
        layout_paned_window.add(layout_text_panel, weight=3)  # 더 큰 비중

        # 줄번호 캔버스와 레이아웃 텍스트
        line_numbers = tk.Canvas(layout_text_panel, width=40, background="#f0f0f0", highlightthickness=0)
        layout_text = tk.Text(layout_text_panel, wrap=tk.NONE, width=50, font=("Arial", 10))
        layout_v_scrollbar = ttk.Scrollbar(layout_text_panel, orient=tk.VERTICAL, command=layout_text.yview)
        layout_h_scrollbar = ttk.Scrollbar(layout_text_panel, orient=tk.HORIZONTAL, command=layout_text.xview)

        # yscrollcommand를 커스텀으로 설정하여 줄번호 갱신 + 스크롤바 연동
        layout_text.configure(
            yscrollcommand=lambda first, last, did=device_id: self.on_layout_yscroll(did, first, last),
            xscrollcommand=layout_h_scrollbar.set
        )

        # 그리드 배치: 줄번호(0), 텍스트(1), 스크롤바(2)
        line_numbers.grid(row=0, column=0, sticky="ns")
        layout_text.grid(row=0, column=1, sticky="nsew")
        layout_v_scrollbar.grid(row=0, column=2, sticky="ns")
        layout_h_scrollbar.grid(row=1, column=1, sticky="ew")

        layout_text_panel.grid_rowconfigure(0, weight=1)
        layout_text_panel.grid_columnconfigure(1, weight=1)

        # 탭 UI 프레임 (PanedWindow의 두 번째 패널)
        tab_ui_frame = ttk.Frame(layout_paned_window)
        layout_paned_window.add(tab_ui_frame, weight=1)  # 더 작은 비중

        # 탭 노트북 생성
        search_notebook = ttk.Notebook(tab_ui_frame)
        search_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Search Items 탭
        search_items_tab = ttk.Frame(search_notebook)
        search_notebook.add(search_items_tab, text="Search Items")

        # 검색 컴포넌트 내부 프레임
        search_content_frame = ttk.Frame(search_items_tab)
        search_content_frame.pack(fill=tk.X, padx=5, pady=5)

        # 왼쪽 프레임 (radio button과 입력 필드들)
        left_search_frame = ttk.Frame(search_content_frame)
        left_search_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 검색 타입 라디오 버튼
        search_type_var = tk.StringVar(value="specific")

        # 첫 번째 라디오 버튼과 입력 필드들
        specific_frame = ttk.Frame(left_search_frame)
        specific_frame.pack(fill=tk.X, pady=(0, 5))

        specific_radio = ttk.Radiobutton(specific_frame, text="", variable=search_type_var, value="specific")
        specific_radio.pack(side=tk.LEFT)

        # 입력 필드들을 위한 그리드 프레임
        fields_frame = ttk.Frame(specific_frame)
        fields_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # text 필드
        ttk.Label(fields_frame, text="text :").grid(row=0, column=0, sticky="w", padx=(0, 5))
        text_entry = ttk.Entry(fields_frame, width=20)
        text_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        # resource-id 필드
        ttk.Label(fields_frame, text="resource-id :").grid(row=0, column=2, sticky="w", padx=(0, 5))
        resource_id_entry = ttk.Entry(fields_frame, width=20)
        resource_id_entry.grid(row=0, column=3, sticky="ew", padx=(0, 10))

        # content-desc/talkback 필드
        ttk.Label(fields_frame, text=f"{var_content_desc} :").grid(row=1, column=0, sticky="w", padx=(0, 5))
        content_desc_entry = ttk.Entry(fields_frame, width=20)
        content_desc_entry.grid(row=1, column=1, sticky="ew", padx=(0, 10))

        # hint 필드
        ttk.Label(fields_frame, text="hint :").grid(row=1, column=2, sticky="w", padx=(0, 5))
        hint_entry = ttk.Entry(fields_frame, width=20)
        hint_entry.grid(row=1, column=3, sticky="ew", padx=(0, 10))

        # package 필드
        ttk.Label(fields_frame, text="package :").grid(row=2, column=0, sticky="w", padx=(0, 5))
        package_entry = ttk.Entry(fields_frame, width=20)
        package_entry.grid(row=2, column=1, sticky="ew", padx=(0, 10))

        # class 필드
        ttk.Label(fields_frame, text="class :").grid(row=2, column=2, sticky="w", padx=(0, 5))
        class_entry = ttk.Entry(fields_frame, width=20)
        class_entry.grid(row=2, column=3, sticky="ew", padx=(0, 10))

        # 그리드 컬럼 가중치 설정
        fields_frame.grid_columnconfigure(1, weight=1)
        fields_frame.grid_columnconfigure(3, weight=1)

        # 두 번째 라디오 버튼과 입력 필드
        all_things_frame = ttk.Frame(left_search_frame)
        all_things_frame.pack(fill=tk.X, pady=(5, 0))

        all_things_radio = ttk.Radiobutton(all_things_frame, text="all things(Regular expression) :", variable=search_type_var, value="all")
        all_things_radio.pack(side=tk.LEFT)

        all_things_entry = ttk.Entry(all_things_frame, width=30)
        all_things_entry.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)

        # 오른쪽 프레임 (검색 버튼)
        right_search_frame = ttk.Frame(search_content_frame)
        right_search_frame.pack(side=tk.RIGHT, padx=(10, 0))

        search_button = ttk.Button(right_search_frame, text="Search",
                                   command=lambda: self.search_in_layout(device_id))
        search_button.pack()

        # 검색 결과 표시 영역
        result_text = tk.Text(search_items_tab, height=20, wrap=tk.WORD, font=("Arial", 10))
        result_scrollbar = ttk.Scrollbar(search_items_tab, orient=tk.VERTICAL, command=result_text.yview)
        result_text.configure(yscrollcommand=result_scrollbar.set)

        result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0), pady=(0, 5))
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 5))

        # GPT 탭
        gpt_tab = ttk.Frame(search_notebook)
        search_notebook.add(gpt_tab, text="GPT")

        # GPT 프롬프트 입력 영역
        gpt_prompt_frame = ttk.LabelFrame(gpt_tab, text="Input Prompt")
        gpt_prompt_frame.pack(fill=tk.X, padx=5, pady=5)

        gpt_prompt_text = scrolledtext.ScrolledText(gpt_prompt_frame, height=5, wrap=tk.WORD)
        gpt_prompt_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # GPT 실행 버튼
        gpt_button_frame = ttk.Frame(gpt_tab)
        gpt_button_frame.pack(fill=tk.X, padx=5, pady=5)

        gpt_run_button = ttk.Button(gpt_button_frame, text="Run",
                                    command=lambda: self.run_gpt_api(device_id))
        gpt_run_button.pack(side=tk.RIGHT)

        # GPT 응답 표시 영역
        gpt_response_frame = ttk.LabelFrame(gpt_tab, text="Response")
        gpt_response_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        gpt_response_text = scrolledtext.ScrolledText(gpt_response_frame, wrap=tk.WORD, font=("Arial", 10))
        gpt_response_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Gemini 탭
        gemini_tab = ttk.Frame(search_notebook)
        search_notebook.add(gemini_tab, text="Gemini")

        # Gemini 프롬프트 입력 영역
        gemini_prompt_frame = ttk.LabelFrame(gemini_tab, text="Input Prompt")
        gemini_prompt_frame.pack(fill=tk.X, padx=5, pady=5)

        gemini_prompt_text = scrolledtext.ScrolledText(gemini_prompt_frame, height=5, wrap=tk.WORD)
        gemini_prompt_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Gemini 실행 버튼
        gemini_button_frame = ttk.Frame(gemini_tab)
        gemini_button_frame.pack(fill=tk.X, padx=5, pady=5)

        gemini_run_button = ttk.Button(gemini_button_frame, text="Run",
                                       command=lambda: self.run_gemini_api(device_id))
        gemini_run_button.pack(side=tk.RIGHT)

        # Gemini 응답 표시 영역
        gemini_response_frame = ttk.LabelFrame(gemini_tab, text="Response")
        gemini_response_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        gemini_response_text = scrolledtext.ScrolledText(gemini_response_frame, wrap=tk.WORD, font=("Arial", 10))
        gemini_response_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 디바이스별 데이터 저장
        self.device_tabs[device_id] = {
            'tab': capture_tab,
            'id_entry': id_entry,
            'screen_canvas': screen_canvas,
            'layout_text': layout_text,
            'current_device_label': current_device_label,
            'font_label': font_label,
            'search_entry': search_entry,
            'screen_image': None,
            'original_image': None,
            'canvas_image': None,
            'selection_start': None,
            'selection_end': None,
            'selection_rect': None,
            'bounds_highlight_rect': None,
            'font_size': 10,
            'undo_stack': [],
            'redo_stack': [],
            'search_positions': [],
            'current_search_index': -1,
            'paned_window': paned_window,
            'coords_label': coords_label,
            'layout_v_scrollbar': layout_v_scrollbar,
            'line_numbers': line_numbers,
            'scale_var': scale_var,  # 축소 비율 변수 추가
            'current_scale': 25,  # 기본 축소 비율 25%
            # 새로운 검색 컴포넌트들
            'search_type_var': search_type_var,
            'text_entry': text_entry,
            'resource_id_entry': resource_id_entry,
            'content_desc_entry': content_desc_entry,
            'hint_entry': hint_entry,
            'package_entry': package_entry,
            'class_entry': class_entry,
            'all_things_entry': all_things_entry,
            'result_text': result_text,
            'layout_paned_window': layout_paned_window,  # splitbar 추가
            # GPT 관련 컴포넌트들
            'gpt_prompt_text': gpt_prompt_text,
            'gpt_response_text': gpt_response_text,
            # Gemini 관련 컴포넌트들
            'gemini_prompt_text': gemini_prompt_text,
            'gemini_response_text': gemini_response_text,
            # 파일 경로 표시 라벨 추가 (요청사항)
            'screen_path_label': screen_path_label,
            'layout_path_label': layout_path_label,
            # 요청사항: x,y,width,height 입력창과 노란색 사각형 관련
            'x_entry': x_entry,
            'y_entry': y_entry,
            'width_entry': width_entry,
            'height_entry': height_entry,
            'yellow_rect': None  # 노란색 dashed 사각형 저장용
        }

        # 마우스 이벤트 바인딩
        # 오른쪽 버튼으로 영역 선택
        screen_canvas.bind('<Button-3>', lambda e: self.on_canvas_right_click(e, device_id))
        screen_canvas.bind('<B3-Motion>', lambda e: self.on_canvas_drag(e, device_id))
        screen_canvas.bind('<ButtonRelease-3>', lambda e: self.on_canvas_right_release(e, device_id))

        # 왼쪽 버튼 클릭: 좌표 표시 + 최소 면적 bounds 선택 (요청사항 2)
        screen_canvas.bind('<Button-1>', lambda e: self.on_canvas_left_click(e, device_id))

        # Screen capture 마우스 휠 이벤트 바인딩
        screen_canvas.bind('<MouseWheel>', lambda e: self.on_screen_scroll(e, device_id))
        screen_canvas.bind('<Button-4>', lambda e: self.on_screen_scroll(e, device_id))  # Linux
        screen_canvas.bind('<Button-5>', lambda e: self.on_screen_scroll(e, device_id))  # Linux

        # Layout text 이벤트 바인딩
        layout_text.bind('<Button-1>', lambda e: self.on_layout_text_click(e, device_id))
        layout_text.bind('<Control-MouseWheel>', lambda e: self.on_font_size_change(e, device_id))
        layout_text.bind('<Control-z>', lambda e: self.undo_text_change(e, device_id))
        layout_text.bind('<Control-y>', lambda e: self.redo_text_change(e, device_id))
        layout_text.bind('<Key>', lambda e: self.on_text_change(e, device_id))
        layout_text.bind('<KeyRelease>', lambda e: self.on_text_change_complete(e, device_id))
        layout_text.bind('<Configure>', lambda e: self.update_line_numbers(device_id))

        # 검색 키바인딩
        layout_text.bind('<F3>', lambda e: self.search_previous(e, device_id))
        layout_text.bind('<F4>', lambda e: self.search_next(e, device_id))
        search_entry.bind('<Return>', lambda e: self.search_text_in_layout(device_id))

        # 우클릭 메뉴 생성
        if not hasattr(self, 'context_menu'):
            self.context_menu = tk.Menu(self.root, tearoff=0)
            self.context_menu.add_command(label="선택영역 저장", command=lambda: self.save_selection(self.current_context_device))
            self.context_menu.add_command(label="복사", command=lambda: self.copy_selection(self.current_context_device))

        # 초기 상태 저장
        self.save_text_state(device_id)

        # 탭을 활성화하고 초기 캡처 실행
        self.notebook.select(capture_tab)
        self.reload_capture(device_id)

        return capture_tab

    def show_yellow_rectangle(self, device_id):
        """입력된 x,y,width,height 값으로 노란색 dashed 사각형 표시"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]

        try:
            # 입력값 가져오기
            x = int(tab_info['x_entry'].get())
            y = int(tab_info['y_entry'].get())
            width = int(tab_info['width_entry'].get())
            height = int(tab_info['height_entry'].get())

            width = width if is_stp_mode else width - x
            height = height if is_stp_mode else height - y


            # 축소 비율 적용
            scale_ratio = tab_info['current_scale'] / 100.0
            scaled_x = int(x * scale_ratio)
            scaled_y = int(y * scale_ratio)
            scaled_width = int(width * scale_ratio)
            scaled_height = int(height * scale_ratio)

            # 기존 노란색 사각형 제거
            if tab_info['yellow_rect']:
                tab_info['screen_canvas'].delete(tab_info['yellow_rect'])

            # 새로운 노란색 dashed 사각형 그리기
            tab_info['yellow_rect'] = tab_info['screen_canvas'].create_rectangle(
                scaled_x, scaled_y,
                scaled_x + scaled_width, scaled_y + scaled_height,
                outline='yellow', dash=(5, 5), width=2
            )

        except ValueError:
            messagebox.showwarning("경고", "x, y, width, height에 올바른 숫자를 입력해주세요.")

    def delete_yellow_rectangle(self, device_id):
        """노란색 dashed 사각형 삭제"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]

        # 노란색 사각형이 있으면 삭제
        if tab_info['yellow_rect']:
            tab_info['screen_canvas'].delete(tab_info['yellow_rect'])
            tab_info['yellow_rect'] = None

    def run_gpt_api(self, device_id):
        """GPT API 호출"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        prompt = tab_info['gpt_prompt_text'].get(1.0, tk.END).strip()

        if not prompt:
            messagebox.showwarning("경고", "프롬프트를 입력해주세요.")
            return

        # 환경변수에서 API 키 가져오기
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            messagebox.showerror("오류", "OPENAI_API_KEY 환경변수가 설정되지 않았습니다.")
            return

        # 응답 영역 초기화
        tab_info['gpt_response_text'].delete(1.0, tk.END)
        tab_info['gpt_response_text'].insert(tk.END, "처리 중...")

        def api_call():
            try:
                url = "https://api.openai.com/v1/chat/completions"

                headers = {
                    "Authorization": f"Bearer {api_key}"
                }

                # 메시지 정의 (prompt 포함)
                data = {
                    "model": "gpt-4.1",
                    "messages": [
                        {"role": "user", "content": prompt}
                    ],
                    "max_tokens": 2000,
                    "temperature": 0.7
                }

                # 업로드할 파일 준비
                files = {
                    "file1": open(r"D:\#python_proj\ScreenLayoutCapture\img\20250915_223829_R3CY207M3MV.png", "rb"),
                    "file2": open(r"D:\#python_proj\ScreenLayoutCapture\layout\20250915_223829_R3CY207M3MV.xml", "rb")
                }

                response = requests.post(
                    url,
                    headers=headers,
                    data={"model": data["model"],
                          "messages": str(data["messages"]),
                          "max_tokens": data["max_tokens"],
                          "temperature": data["temperature"]},
                    files=files,
                    timeout=30
                )

                # 파일 닫기
                files["file1"].close()
                files["file2"].close()

                if response.status_code == 200:
                    result = response.json()
                    answer = result["choices"][0]["message"]["content"]
                    self.root.after(0, lambda: self.update_gpt_response(device_id, answer))
                else:
                    error_msg = f"API 호출 실패: {response.status_code}\n{response.text}"
                    self.root.after(0, lambda: self.update_gpt_response(device_id, error_msg))

            except requests.exceptions.RequestException as e:
                error_msg = f"네트워크 오류: {str(e)}"
                self.root.after(0, lambda: self.update_gpt_response(device_id, error_msg))
            except Exception as e:
                error_msg = f"오류 발생: {str(e)}"
                self.root.after(0, lambda: self.update_gpt_response(device_id, error_msg))

        # 별도 스레드에서 API 호출
        thread = threading.Thread(target=api_call, daemon=True)
        thread.start()

    def update_gpt_response(self, device_id, response):
        """GPT 응답 UI 업데이트"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        tab_info['gpt_response_text'].delete(1.0, tk.END)
        tab_info['gpt_response_text'].insert(1.0, response)

    def run_gemini_api(self, device_id):
        """Gemini API 호출"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        prompt = tab_info['gemini_prompt_text'].get(1.0, tk.END).strip()

        if not prompt:
            messagebox.showwarning("경고", "프롬프트를 입력해주세요.")
            return

        # 환경변수에서 API 키 가져오기
        api_key = os.getenv('GOOGLE_GEMINI_KEY')
        if not api_key:
            messagebox.showerror("오류", "GOOGLE_GEMINI_KEY 환경변수가 설정되지 않았습니다.")
            return

        # 응답 영역 초기화
        tab_info['gemini_response_text'].delete(1.0, tk.END)
        tab_info['gemini_response_text'].insert(tk.END, "처리 중...")

        def api_call():
            try:
                url = f'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5:generateContent?key={api_key}'

                data = {
                    'contents': [{
                        'parts': [{
                            'text': prompt
                        }]
                    }]
                }

                response = requests.post(
                    url,
                    headers={'Content-Type': 'application/json'},
                    json=data,
                    timeout=30
                )

                if response.status_code == 200:
                    result = response.json()
                    if 'candidates' in result and len(result['candidates']) > 0:
                        answer = result['candidates'][0]['content']['parts'][0]['text']
                    else:
                        answer = "응답을 받을 수 없습니다."

                    # UI 업데이트는 메인 스레드에서 실행
                    self.root.after(0, lambda: self.update_gemini_response(device_id, answer))
                else:
                    error_msg = f"API 호출 실패: {response.status_code}\n{response.text}"
                    self.root.after(0, lambda: self.update_gemini_response(device_id, error_msg))

            except requests.exceptions.RequestException as e:
                error_msg = f"네트워크 오류: {str(e)}"
                self.root.after(0, lambda: self.update_gemini_response(device_id, error_msg))
            except Exception as e:
                error_msg = f"오류 발생: {str(e)}"
                self.root.after(0, lambda: self.update_gemini_response(device_id, error_msg))

        # 별도 스레드에서 API 호출
        thread = threading.Thread(target=api_call, daemon=True)
        thread.start()

    def update_gemini_response(self, device_id, response):
        """Gemini 응답 UI 업데이트"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        tab_info['gemini_response_text'].delete(1.0, tk.END)
        tab_info['gemini_response_text'].insert(1.0, response)

    def search_in_layout(self, device_id):
        """새로운 검색 기능"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        search_type = tab_info['search_type_var'].get()

        # 레이아웃 텍스트 내용 가져오기
        layout_content = tab_info['layout_text'].get(1.0, tk.END)

        # 이전 검색 하이라이트 제거
        tab_info['layout_text'].tag_remove("search_highlight", 1.0, tk.END)
        tab_info['result_text'].delete(1.0, tk.END)

        matching_lines = []

        if search_type == "specific":
            # 특정 필드 검색
            text_val = tab_info['text_entry'].get().strip()
            resource_id_val = tab_info['resource_id_entry'].get().strip()
            content_desc_val = tab_info['content_desc_entry'].get().strip()
            hint_val = tab_info['hint_entry'].get().strip()
            package_val = tab_info['package_entry'].get().strip()
            class_val = tab_info['class_entry'].get().strip()

            # 최소 한 개의 필드에는 값이 있어야 함
            search_fields = [text_val, resource_id_val, content_desc_val, hint_val, package_val, class_val]
            if not any(search_fields):
                messagebox.showwarning("경고", "검색할 값을 하나 이상 입력해주세요.")
                return

            lines = layout_content.split('\n')
            for line_num, line in enumerate(lines, 1):
                if self.matches_specific_search(line, text_val, resource_id_val, content_desc_val,
                                                hint_val, package_val, class_val):
                    matching_lines.append((line_num, line.strip()))

        elif search_type == "all":
            # 전체 속성 검색 (Regular expression으로 변경)
            all_things_val = tab_info['all_things_entry'].get().strip()
            if not all_things_val:
                messagebox.showwarning("경고", "검색할 값을 입력해주세요.")
                return

            lines = layout_content.split('\n')
            for line_num, line in enumerate(lines, 1):
                if self.matches_all_things_search_regex(line, all_things_val):
                    matching_lines.append((line_num, line.strip()))

        # 결과 표시
        if matching_lines:
            # 레이아웃 텍스트에서 해당 라인들 하이라이트
            for line_num, line_content in matching_lines:
                line_start = f"{line_num}.0"
                line_end = f"{line_num}.end"
                tab_info['layout_text'].tag_add("search_highlight", line_start, line_end)

            # 검색 하이라이트 스타일 설정
            tab_info['layout_text'].tag_configure("search_highlight", background="yellow")

            # 첫 번째 결과로 스크롤
            first_line = matching_lines[0][0]
            tab_info['layout_text'].see(f"{first_line}.0")

            # 결과 텍스트에 표시 (요청사항 반영: x,y,width,height 값과 전체 raw 문자열 표시)
            result_text = f""
            for line_num, line_content in matching_lines:
                # bounds 정보 추출
                # for stp_mode
                if is_stp_mode:
                    bounds_match = re.search(r'bounds="\{(\d+), (\d+), (\d+), (\d+)\}"', line_content)
                else:
                    bounds_match = re.search(r'bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', line_content)

                if bounds_match:
                    x1, y1, x2, y2 = bounds_match.groups()
                    width = int(x2) if is_stp_mode else int(x2) - int(x1)
                    height = int(y2) if is_stp_mode else int(y2) - int(y1)
                    result_text += f"x={x1},y={y1},w={width},h={height}: Line {line_num}: {line_content}\n"
                else:
                    result_text += f"Line {line_num}: {line_content}\n"

            tab_info['result_text'].insert(1.0, result_text)

        else:
            tab_info['result_text'].insert(1.0, "")

    def matches_specific_search(self, line, text_val, resource_id_val, content_desc_val,
                                hint_val, package_val, class_val):
        """특정 필드 검색 매칭 로직"""
        # 값이 있는 필드들만 체크
        checks = []

        if text_val:
            checks.append(f'text="{text_val}"' in line)
        if resource_id_val:
            checks.append(f'resource-id="{resource_id_val}"' in line)
        if content_desc_val:
            checks.append(f'{var_content_desc}="{content_desc_val}"' in line)
        if hint_val:
            checks.append(f'hint="{hint_val}"' in line)
        if package_val:
            checks.append(f'package="{package_val}"' in line)
        if class_val:
            checks.append(f'class="{class_val}"' in line)

        # 모든 조건이 만족되어야 함
        return len(checks) > 0 and all(checks)

    def matches_all_things_search_regex(self, line, search_val):
        """전체 속성 검색 매칭 로직 (Regular expression 방식으로 변경)"""
        try:
            # 정규표현식 패턴 컴파일
            pattern = re.compile(search_val, re.IGNORECASE)

            # 모든 속성에서 검색 (정규표현식으로 매칭)
            for attr in g_attributes:
                attr_start = line.find(attr)
                if attr_start != -1:
                    # 속성 값 추출 (따옴표 사이의 값)
                    quote_start = line.find('"', attr_start)
                    if quote_start != -1:
                        quote_end = line.find('"', quote_start + 1)
                        if quote_end != -1:
                            attr_value = line[quote_start + 1:quote_end]
                            # 정규표현식으로 매칭 검사
                            if pattern.search(attr_value):
                                return True
            return False
        except re.error:
            # 잘못된 정규표현식인 경우 일반 문자열 검색으로 fallback
            return self.matches_all_things_search(line, search_val)

    def matches_all_things_search(self, line, search_val):
        """전체 속성 검색 매칭 로직 (기존 문자열 일치 방식)"""
        # 모든 속성에서 검색 값이 포함되어 있는지 확인
        for attr in g_attributes:
            attr_start = line.find(attr)
            if attr_start != -1:
                # 속성 값 추출 (따옴표 사이의 값)
                quote_start = line.find('"', attr_start)
                if quote_start != -1:
                    quote_end = line.find('"', quote_start + 1)
                    if quote_end != -1:
                        attr_value = line[quote_start + 1:quote_end]
                        if search_val in attr_value:
                            return True
        return False

    def is_click_on_image(self, event, device_id):
        """마우스 클릭이 이미지 영역 내에 있는지 확인"""
        if device_id not in self.device_tabs:
            return False

        tab_info = self.device_tabs[device_id]
        if not tab_info['original_image'] or not tab_info['canvas_image']:
            return False

        screen_canvas = tab_info['screen_canvas']

        # 스크롤을 고려한 실제 캔버스 좌표
        click_x = int(screen_canvas.canvasx(event.x))
        click_y = int(screen_canvas.canvasy(event.y))

        # 이미지의 실제 크기 (축소 비율 적용된 크기)
        scale_ratio = tab_info['current_scale'] / 100.0
        image_width = int(tab_info['original_image'].width * scale_ratio)
        image_height = int(tab_info['original_image'].height * scale_ratio)

        # 이미지는 (0, 0)에서 시작하므로, 클릭 좌표가 이미지 범위 내에 있는지 확인
        return 0 <= click_x <= image_width and 0 <= click_y <= image_height

    def on_scale_change(self, device_id):
        """축소 비율 변경 시 이미지 다시 표시 (요청사항 1)"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        tab_info['current_scale'] = int(tab_info['scale_var'].get())
        self.display_image(device_id)

    def on_alt_f4(self, event):
        """Alt+F4 키 처리 (요청사항 3)"""
        if event.state & 0x4:  # Ctrl 키가 눌려있는 경우
            self.root.destroy()  # 프로그램 종료
        else:
            self.minimize_to_tray()
            return "break"

    def on_layout_yscroll(self, device_id, first, last):
        """레이아웃 텍스트 yscroll 동기화 + 줄번호 갱신"""
        if device_id not in self.device_tabs:
            return
        tab_info = self.device_tabs[device_id]
        # 실제 스크롤바 위치 설정
        tab_info['layout_v_scrollbar'].set(first, last)
        # 줄번호 갱신
        self.update_line_numbers(device_id)

    def update_line_numbers(self, device_id):
        """레이아웃 텍스트의 줄번호 표시 갱신"""
        if device_id not in self.device_tabs:
            return
        tab_info = self.device_tabs[device_id]
        layout_text = tab_info['layout_text']
        line_numbers = tab_info['line_numbers']

        # 전체 라인 수
        total_lines = int(layout_text.index('end-1c').split('.')[0])
        # 폰트 정보
        font = tkfont.Font(font=layout_text['font'])
        digits = max(2, len(str(total_lines)))
        width_px = font.measure('9' * digits) + 8
        line_numbers.config(width=width_px)

        # 현재 보이는 첫/마지막 라인
        first_visible_line = int(layout_text.index("@0,0").split('.')[0])
        last_visible_line = int(layout_text.index(f"@0,{layout_text.winfo_height()}").split('.')[0])

        # 기준 y 오프셋
        first_info = layout_text.dlineinfo(f"{first_visible_line}.0")
        base_y = first_info[1] if first_info else 0

        # 캔버스 클리어
        line_numbers.delete("all")

        # 줄번호 그리기
        for line in range(first_visible_line, last_visible_line + 1):
            info = layout_text.dlineinfo(f"{line}.0")
            if info:
                y = info[1] - base_y
                number = str(line)
                line_numbers.create_text(
                    width_px - 4, y,
                    anchor="ne",
                    text=number,
                    font=layout_text['font'],
                    fill="#555555"
                )

    def on_canvas_right_click(self, event, device_id):
        """캔버스 오른쪽 클릭 시작 (영역 선택용)"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        canvas = tab_info['screen_canvas']
        x = int(canvas.canvasx(event.x))
        y = int(canvas.canvasy(event.y))
        tab_info['selection_start'] = (x, y)
        if tab_info['selection_rect']:
            tab_info['screen_canvas'].delete(tab_info['selection_rect'])
            tab_info['selection_rect'] = None

    def on_canvas_right_release(self, event, device_id):
        """캔버스 오른쪽 릴리즈 (영역 선택 완료 후 컨텍스트 메뉴)"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        if tab_info['selection_start']:
            canvas = tab_info['screen_canvas']
            x = int(canvas.canvasx(event.x))
            y = int(canvas.canvasy(event.y))
            tab_info['selection_end'] = (x, y)

            # 선택 영역이 있으면 컨텍스트 메뉴 표시
            if tab_info['selection_rect']:
                self.current_context_device = device_id
                self.context_menu.post(event.x_root, event.y_root)

    def on_canvas_left_click(self, event, device_id):
        """캔버스 왼쪽 클릭: 좌표 표시 + 최소 면적 bounds 아이템 라인 선택 (요청사항 2)"""
        if device_id not in self.device_tabs:
            return

        # 이미지 영역 내에서만 동작하도록 체크
        if not self.is_click_on_image(event, device_id):
            return

        tab_info = self.device_tabs[device_id]
        if not tab_info['original_image']:
            return

        screen_canvas = tab_info['screen_canvas']
        # 스크롤을 고려한 실제 캔버스 좌표
        click_x = int(screen_canvas.canvasx(event.x))
        click_y = int(screen_canvas.canvasy(event.y))

        # 축소 비율을 고려하여 100% 기준 좌표로 변환 (요청사항 1)
        scale_ratio = tab_info['current_scale'] / 100.0
        actual_x = int(click_x / scale_ratio)
        actual_y = int(click_y / scale_ratio)

        # 좌표 라벨 업데이트 (100% 기준 좌표로 표시) - 요청사항 1: 이미지 Resolution 추가
        if tab_info.get('coords_label') and tab_info['original_image']:
            resolution = f"{tab_info['original_image'].width}x{tab_info['original_image'].height}"
            tab_info['coords_label'].configure(text=f"이미지 Resolution : {resolution}, 좌표 : ({actual_x},{actual_y})")

        # Layout XML에서 해당 좌표를 포함하는 bounds 중 최소 면적 라인 찾기
        layout_content = tab_info['layout_text'].get(1.0, tk.END)
        result = self.find_smallest_bounds_line(layout_content, actual_x, actual_y)

        if result:
            line_num, bounds = result
            self.select_line_in_layout(device_id, line_num)
            # 선택된 라인의 bounds를 화면에 표시 (축소 비율 적용)
            if bounds:
                x1, y1, x2, y2 = bounds
                scaled_x1 = int(x1 * scale_ratio)
                scaled_y1 = int(y1 * scale_ratio)
                scaled_x2 = int(x2 * scale_ratio)
                scaled_y2 = int(y2 * scale_ratio)
                self.highlight_bounds_on_screen(device_id, scaled_x1, scaled_y1, scaled_x2, scaled_y2)

    def find_smallest_bounds_line(self, xml_content, x, y):
        """
        XML 내용에서 주어진 좌표를 포함하는 bounds를 가진 라인들 중
        width*height 면적이 가장 작은 라인의 번호와 bounds를 반환.
        """
        lines = xml_content.split('\n')
        smallest_area = None
        best_line_num = None
        best_bounds = None

        # bounds="[\d+, \d+, \d+, \d+" 패턴
        # for stp_mode
        if is_stp_mode:
            pattern = re.compile(r'bounds="\{(\d+), (\d+), (\d+), (\d+)\}"')
        else:
            pattern = re.compile(r'bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"')


        for line_num, line in enumerate(lines, start=1):
            m = pattern.search(line)
            if not m:
                continue
            x1, y1, x2, y2 = map(int, m.groups())

            # for stp_mode
            x2 = x1 + x2 if is_stp_mode else x2
            y2 = y1 + y2 if is_stp_mode else y2

            if x1 <= x <= x2 and y1 <= y <= y2:
                area = max(0, x2 - x1) * max(0, y2 - y1)
                if smallest_area is None or area < smallest_area:
                    smallest_area = area
                    best_line_num = line_num
                    best_bounds = (x1, y1, x2, y2)

        if best_line_num is not None:
            return best_line_num, best_bounds
        return None

    def select_line_in_layout(self, device_id, line_number):
        """Layout text에서 특정 라인을 선택"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        layout_text = tab_info['layout_text']

        # 라인 위치로 이동
        line_start = f"{line_number}.0"
        line_end = f"{line_number}.end"

        # 해당 라인 선택
        layout_text.tag_remove(tk.SEL, 1.0, tk.END)
        layout_text.tag_add(tk.SEL, line_start, line_end)
        layout_text.mark_set(tk.INSERT, line_start)
        layout_text.see(line_start)

    def on_layout_text_click(self, event, device_id):
        """Layout text 클릭 시 해당 라인의 bounds를 Screen capture에 표시"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        layout_text = tab_info['layout_text']

        # 클릭된 위치의 라인 번호 가져오기
        line_index = layout_text.index(f"@{event.x},{event.y}")
        line_num = int(line_index.split('.')[0])

        # 해당 라인의 내용 가져오기
        line_start = f"{line_num}.0"
        line_end = f"{line_num}.end"
        line_content = layout_text.get(line_start, line_end)

        # bounds 속성 찾기
        # for stp_mode
        if is_stp_mode:
            m = re.search(r'bounds="\{(\d+), (\d+), (\d+), (\d+)\}"', line_content)
        else:
            m = re.search(r'bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', line_content)

        if m:
            x1, y1, x2, y2 = map(int, m.groups())

            # for stp_mode
            x2 = x1 + x2 if is_stp_mode else x2
            y2 = y1 + y2 if is_stp_mode else y2

            # 축소 비율 적용
            scale_ratio = tab_info['current_scale'] / 100.0
            scaled_x1 = int(x1 * scale_ratio)
            scaled_y1 = int(y1 * scale_ratio)
            scaled_x2 = int(x2 * scale_ratio)
            scaled_y2 = int(y2 * scale_ratio)
            # Screen capture에 파란색 dashed 사각형 표시
            self.highlight_bounds_on_screen(device_id, scaled_x1, scaled_y1, scaled_x2, scaled_y2)

    def highlight_bounds_on_screen(self, device_id, x1, y1, x2, y2):
        """Screen capture에서 bounds 영역을 파란색 dashed line으로 표시"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        screen_canvas = tab_info['screen_canvas']

        # 기존 하이라이트 제거
        if tab_info['bounds_highlight_rect']:
            screen_canvas.delete(tab_info['bounds_highlight_rect'])

        # 새로운 하이라이트 사각형 그리기 (파란색 dashed line)
        tab_info['bounds_highlight_rect'] = screen_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline='blue', dash=(5, 5), width=2
        )

    def on_screen_scroll(self, event, device_id):
        """Screen capture 영역에서 마우스 휠 스크롤 처리"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        screen_canvas = tab_info['screen_canvas']

        # 현재 스크롤 위치 가져오기
        top, bottom = screen_canvas.yview()

        # 스크롤 델타 계산 (Windows)
        if hasattr(event, 'delta') and event.delta != 0:
            delta = -1 * (event.delta / 120)
        # Linux의 경우
        elif hasattr(event, 'num') and event.num == 4:
            delta = -1
        elif hasattr(event, 'num') and event.num == 5:
            delta = 1
        else:
            return

        # 스크롤 실행 (세로 스크롤 우선)
        screen_canvas.yview_scroll(int(delta), "units")

    def on_tab_right_click(self, event, tab):
        """탭 우클릭 시 종료 메뉴 표시"""
        if len(self.notebook.tabs()) > 1:  # Device Management 탭은 유지
            tab_menu = tk.Menu(self.root, tearoff=0)
            tab_menu.add_command(label="탭 닫기", command=lambda: self.close_tab(tab))
            tab_menu.post(event.x_root, event.y_root)

    def close_tab(self, tab):
        """탭 닫기"""
        # 해당 디바이스 ID 찾기
        device_to_remove = None
        for device_id, tab_info in self.device_tabs.items():
            if tab_info['tab'] == tab:
                device_to_remove = device_id
                break

        if device_to_remove:
            # 탭 삭제
            self.notebook.forget(tab)
            # 디바이스 정보 삭제
            del self.device_tabs[device_to_remove]

    def on_font_size_change(self, event, device_id):
        """Ctrl + 마우스휠로 폰트 크기 변경"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]

        if event.delta > 0 and tab_info['font_size'] < 50:
            tab_info['font_size'] += 1
        elif event.delta < 0 and tab_info['font_size'] > 3:
            tab_info['font_size'] -= 1

        tab_info['layout_text'].configure(font=("Arial", tab_info['font_size']))
        tab_info['font_label'].configure(text=f"Font : {tab_info['font_size']}")

        # 줄번호 갱신
        self.update_line_numbers(device_id)

    def on_text_change(self, event, device_id):
        """텍스트 변경 시작"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        if not tab_info.get('is_recording_change', False) and event.keysym not in ['Control_L', 'Control_R', 'z', 'y']:
            tab_info['is_recording_change'] = True

    def on_text_change_complete(self, event, device_id):
        """텍스트 변경 완료"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        if tab_info.get('is_recording_change', False) and event.keysym not in ['Control_L', 'Control_R', 'z', 'y']:
            self.save_text_state(device_id)
            tab_info['is_recording_change'] = False
            # 줄번호 갱신
            self.update_line_numbers(device_id)

    def save_text_state(self, device_id):
        """현재 텍스트 상태 저장"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        current_text = tab_info['layout_text'].get(1.0, tk.END)

        if not tab_info['undo_stack'] or tab_info['undo_stack'][-1] != current_text:
            tab_info['undo_stack'].append(current_text)
            if len(tab_info['undo_stack']) > 100:
                tab_info['undo_stack'].pop(0)
            tab_info['redo_stack'].clear()

    def undo_text_change(self, event, device_id):
        """텍스트 변경 취소 (Ctrl+Z)"""
        if device_id not in self.device_tabs:
            return "break"

        tab_info = self.device_tabs[device_id]

        if len(tab_info['undo_stack']) > 1:
            current_text = tab_info['undo_stack'].pop()
            tab_info['redo_stack'].append(current_text)
            previous_text = tab_info['undo_stack'][-1]

            tab_info['layout_text'].delete(1.0, tk.END)
            tab_info['layout_text'].insert(1.0, previous_text)
            # 줄번호 갱신
            self.update_line_numbers(device_id)
        return "break"

    def redo_text_change(self, event, device_id):
        """텍스트 변경 다시실행 (Ctrl+Y)"""
        if device_id not in self.device_tabs:
            return "break"

        tab_info = self.device_tabs[device_id]

        if tab_info['redo_stack']:
            next_text = tab_info['redo_stack'].pop()
            tab_info['undo_stack'].append(next_text)

            tab_info['layout_text'].delete(1.0, tk.END)
            tab_info['layout_text'].insert(1.0, next_text)
            # 줄번호 갱신
            self.update_line_numbers(device_id)
        return "break"

    def search_text_in_layout(self, device_id, event=None):
        """레이아웃 텍스트에서 검색"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        search_term = tab_info['search_entry'].get().strip()
        if not search_term:
            return

        content = tab_info['layout_text'].get(1.0, tk.END)

        # 검색 결과 초기화
        tab_info['search_positions'] = []
        tab_info['current_search_index'] = -1

        # 이전 검색 하이라이트 제거
        tab_info['layout_text'].tag_remove("search_highlight", 1.0, tk.END)

        # 검색 실행
        start = 1.0
        while True:
            pos = tab_info['layout_text'].search(search_term, start, tk.END)
            if not pos:
                break

            end_pos = f"{pos}+{len(search_term)}c"
            tab_info['search_positions'].append((pos, end_pos))
            tab_info['layout_text'].tag_add("search_highlight", pos, end_pos)
            start = end_pos

        # 검색 하이라이트 스타일 설정
        tab_info['layout_text'].tag_configure("search_highlight", background="yellow")

        if tab_info['search_positions']:
            tab_info['current_search_index'] = 0
            self.highlight_current_search(device_id)
            messagebox.showinfo("검색 결과", f"{len(tab_info['search_positions'])}개 항목을 찾았습니다.")
        else:
            messagebox.showinfo("검색 결과", "검색 결과가 없습니다.")

    def search_next(self, event, device_id):
        """다음 검색 결과로 이동 (F4)"""
        if device_id not in self.device_tabs:
            return "break"

        tab_info = self.device_tabs[device_id]

        if not tab_info['search_positions']:
            return "break"

        if tab_info['current_search_index'] < len(tab_info['search_positions']) - 1:
            tab_info['current_search_index'] += 1
            self.highlight_current_search(device_id)
        else:
            messagebox.showinfo("검색", "마지막 검색 결과입니다.")

        return "break"

    def search_previous(self, event, device_id):
        """이전 검색 결과로 이동 (F3)"""
        if device_id not in self.device_tabs:
            return "break"

        tab_info = self.device_tabs[device_id]

        if not tab_info['search_positions']:
            return "break"

        if tab_info['current_search_index'] > 0:
            tab_info['current_search_index'] -= 1
            self.highlight_current_search(device_id)
        else:
            messagebox.showinfo("검색", "첫 번째 검색 결과입니다.")

        return "break"

    def highlight_current_search(self, device_id):
        """현재 검색 결과 하이라이트 및 스크롤"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]

        if 0 <= tab_info['current_search_index'] < len(tab_info['search_positions']):
            # 이전 현재 하이라이트 제거
            tab_info['layout_text'].tag_remove("current_search", 1.0, tk.END)

            # 현재 검색 결과 하이라이트
            pos, end_pos = tab_info['search_positions'][tab_info['current_search_index']]
            tab_info['layout_text'].tag_add("current_search", pos, end_pos)
            tab_info['layout_text'].tag_configure("current_search", background="orange", foreground="black")

            # 해당 위치로 스크롤
            tab_info['layout_text'].see(pos)

    def on_ctrl_press(self, event):
        self.ctrl_pressed = True

    def on_ctrl_release(self, event):
        self.ctrl_pressed = False

    def on_closing(self):
        """창 닫기 처리 - 시스템 트레이로 이동 (요청사항 3)"""
        if self.ctrl_pressed:
            if self.tray_icon:
                try:
                    self.tray_icon.visible = False
                except Exception:
                    pass

                try:
                    self.tray_icon.stop()
                except Exception:
                    pass

            try:
                self.root.quit()
            finally:
                try:
                    self.root.destroy()
                except Exception:
                    pass
        else:
            self.minimize_to_tray()

    def minimize_to_tray(self):
        """시스템 트레이로 최소화"""
        self.root.withdraw()
        self.is_minimized_to_tray = True

        # 트레이 아이콘 생성
        if not self.tray_icon:
            # 아이콘 파일 경로
            icon_path = os.path.join(os.path.dirname(__file__), "icon_ScreenLayoutCapture.png")
            try:
                image = Image.open(icon_path)
            except Exception as e:
                # 아이콘 파일이 없을 경우 예외처리(기본 파랑 정사각형 사용)
                image = Image.new('RGB', (64, 64), color='blue')
                draw = ImageDraw.Draw(image)
                draw.rectangle([16, 16, 48, 48], fill='white')

            menu = pystray.Menu(
                item('열기', self.restore_from_tray),
                item('종료', self.quit_from_tray)
            )

            self.tray_icon = pystray.Icon("ScreenCapture", image, "Screen & Layout Capture", menu)

            # 트레이 아이콘을 별도 스레드에서 실행
            tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
            tray_thread.start()

    def restore_from_tray(self, icon=None, item=None):
        """트레이에서 복원"""
        self.is_minimized_to_tray = False
        self.root.deiconify()
        self.root.lift()
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None

    def quit_from_tray(self, icon=None, item=None):
        """트레이에서 종료"""
        if self.tray_icon:
            try:
                self.tray_icon.visible = False
            except Exception:
                pass

            try:
                self.tray_icon.stop()
            except Exception:
                pass

        try:
            self.root.quit()
        finally:
            try:
                self.root.destroy()
            except Exception:
                pass

    def load_devices(self):
        """연결된 디바이스 목록 로드"""
        try:
            result = subprocess.run(['adb', 'devices'], capture_output=True, text=True, check=True)
            lines = result.stdout.strip().split('\n')[1:]  # 첫 번째 헤더 라인 제외

            self.device_listbox.delete(0, tk.END)
            for line in lines:
                if line.strip() and '\tdevice' in line:
                    device_id = line.split('\t')[0]
                    self.device_listbox.insert(tk.END, device_id)
        except subprocess.CalledProcessError:
            messagebox.showerror("오류", "ADB를 찾을 수 없거나 디바이스를 읽을 수 없습니다.")
        except FileNotFoundError:
            messagebox.showerror("오류", "ADB가 설치되어 있지 않거나 PATH에 없습니다.")

    def on_device_select(self, event):
        """디바이스 선택 이벤트"""
        selection = self.device_listbox.curselection()
        if selection:
            self.current_device = self.device_listbox.get(selection[0])
            self.get_device_info()

    def on_device_double_click(self, event):
        """디바이스 더블클릭 이벤트 - 새 탭 생성"""
        selection = self.device_listbox.curselection()
        if selection:
            device_id = self.device_listbox.get(selection[0])

            # 이미 해당 디바이스의 탭이 존재하는지 확인
            if device_id not in self.device_tabs:
                self.create_capture_tab(device_id)
            else:
                # 기존 탭이 있으면 해당 탭으로 이동하고 reload cap 실행
                self.notebook.select(self.device_tabs[device_id]['tab'])
                self.reload_capture(device_id)

    def get_device_info(self):
        """선택된 디바이스의 상세 정보 가져오기"""
        if not self.current_device:
            return

        try:
            # 디바이스 속성 정보 가져오기
            result = subprocess.run(['adb', '-s', self.current_device, 'shell', 'getprop'],
                                    capture_output=True, text=True, check=True, encoding='utf-8')

            # 주요 정보만 필터링
            important_props = [
                'ro.product.model', 'ro.product.brand', 'ro.product.manufacturer',
                'ro.build.version.release', 'ro.build.version.sdk',
                'ro.product.cpu.abi', 'ro.build.display.id'
            ]

            info_text = f"Device ID: {self.current_device}\n\n"

            for line in result.stdout.split('\n'):
                for prop in important_props:
                    if f'[{prop}]' in line:
                        value = line.split(']: [')[1].rstrip(']') if ']: [' in line else 'Unknown'
                        prop_name = prop.replace('ro.product.', '').replace('ro.build.', '').replace('.', ' ').title()
                        info_text += f"{prop_name}: {value}\n"

            # 요청사항 2: mViewports 정보 추가
            try:
                display_result = subprocess.run(['adb', '-s', self.current_device, 'shell', 'dumpsys', 'display'],
                                                capture_output=True, text=True, check=True, encoding='utf-8')

                # mViewports 정보 추출
                viewport_info = self.extract_viewport_info(display_result.stdout)
                if viewport_info:
                    info_text += f"\n{viewport_info}"

            except subprocess.CalledProcessError:
                pass  # dumpsys display 실패 시 무시

            self.device_info_text.delete(1.0, tk.END)
            self.device_info_text.insert(1.0, info_text)

        except subprocess.CalledProcessError as e:
            self.device_info_text.delete(1.0, tk.END)
            self.device_info_text.insert(1.0, f"디바이스 정보를 가져올 수 없습니다: {e}")

    def extract_viewport_info(self, dumpsys_output):
        """dumpsys display 출력에서 mViewports 정보 추출 (요청사항 2)"""
        try:
            # mViewports 라인 찾기
            for line in dumpsys_output.split('\n'):
                if 'mViewports=' in line:
                    # DisplayViewport 패턴 찾기
                    pattern = r'DisplayViewport\{[^}]*displayId=(\d+)[^}]*deviceWidth=(\d+)[^}]*deviceHeight=(\d+)[^}]*\}'
                    matches = re.findall(pattern, line)

                    if matches:
                        result = ""
                        for match in matches:
                            display_id, width, height = match
                            if display_id == '0':
                                result += f"main display : {width}x{height}\n"
                            else:
                                result += f"sub display : {width}x{height}\n"
                        return result.strip()
            return None
        except Exception:
            return None

    def reload_capture(self, device_id):
        """화면 및 레이아웃 캡처 새로고침"""
        if device_id not in self.device_tabs:
            return

        # ID 업데이트
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        current_id = f"{timestamp}_{device_id}"

        tab_info = self.device_tabs[device_id]
        tab_info['id_entry'].delete(0, tk.END)
        tab_info['id_entry'].insert(0, current_id)

        # 화면 캡처
        self.capture_screen(device_id)

        # 레이아웃 캡처
        self.capture_layout(device_id)

    def capture_screen(self, device_id):
        """화면 캡처"""
        if device_id not in self.device_tabs:
            return

        try:
            # ADB로 스크린샷 캡처
            subprocess.run(['adb', '-s', device_id, 'shell', 'screencap', '/sdcard/screenshot.png'], check=True)
            subprocess.run(['adb', '-s', device_id, 'pull', '/sdcard/screenshot.png', f'temp_screenshot_{device_id}.png'], check=True)

            # 이미지 로드 및 표시
            original_image = Image.open(f'temp_screenshot_{device_id}.png')
            self.device_tabs[device_id]['original_image'] = original_image
            self.display_image(device_id)

            # 임시 파일 삭제
            os.remove(f'temp_screenshot_{device_id}.png')

        except subprocess.CalledProcessError as e:
            messagebox.showerror("오류", f"화면 캡처 실패: {e}")

    def display_image(self, device_id):
        """캔버스에 이미지 표시 (축소 비율 적용, 요청사항 1)"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        if not tab_info['original_image']:
            return

        original_image = tab_info['original_image']
        screen_canvas = tab_info['screen_canvas']

        # 축소 비율 적용 (요청사항 1)
        scale_ratio = tab_info['current_scale'] / 100.0
        new_width = int(original_image.width * scale_ratio)
        new_height = int(original_image.height * scale_ratio)

        # 이미지 리사이즈
        resized_image = original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        screen_image = ImageTk.PhotoImage(resized_image)
        tab_info['screen_image'] = screen_image

        # 캔버스 클리어 후 이미지 표시 (왼쪽 정렬 - x=0)
        screen_canvas.delete("all")
        canvas_image = screen_canvas.create_image(0, 0, anchor=tk.NW, image=screen_image)
        tab_info['canvas_image'] = canvas_image

        # 스크롤 영역 설정 (리사이즈된 이미지 크기로 설정)
        screen_canvas.configure(scrollregion=(0, 0, new_width, new_height))

        # 요청사항 1: 좌표 라벨에 이미지 Resolution 표시
        if tab_info.get('coords_label'):
            resolution = f"{original_image.width}x{original_image.height}"
            tab_info['coords_label'].configure(text=f"이미지 Resolution : {resolution}, 좌표 : (-,-)")

    def capture_layout(self, device_id):
        """레이아웃 캡처 (UI Automator dump)"""
        if device_id not in self.device_tabs:
            return

        try:
            # uiautomator로 덤프 생성하여 /sdcard/layout.xml에 저장
            subprocess.run(['adb', '-s', device_id, 'shell', 'uiautomator', 'dump', '/sdcard/layout.xml'],
                           check=True, capture_output=True, text=True)

            # 충분한 대기 시간 확보 (파일 생성 및 저장 완료 대기)
            time.sleep(3)

            # adb pull 명령어를 통해 PC로 파일 가져오기
            subprocess.run(['adb', '-s', device_id, 'pull', '/sdcard/layout.xml', f'temp_layout_{device_id}.xml'],
                           check=True)

            # PC로 가져온 파일 읽기
            # for stp_mode
            tmp_file_name = 'temp.xml' if is_stp_mode else f'temp_layout_{device_id}.xml'

            with open(tmp_file_name, 'r', encoding='utf-8') as f:
                xml_content = f.read()

            # XML을 pretty print 형태로 변환
            root = ET.fromstring(xml_content)
            rough_string = ET.tostring(root, 'unicode')
            reparsed = xml.dom.minidom.parseString(rough_string)
            layout_xml = reparsed.toprettyxml(indent="  ")

            # for stp_mode
            if is_stp_mode:
                layout_xml = layout_xml.replace('\r\n', '').replace('\r', '')

            # 텍스트 위젯에 표시
            tab_info = self.device_tabs[device_id]
            tab_info['layout_text'].delete(1.0, tk.END)
            tab_info['layout_text'].insert(1.0, layout_xml)

            # Undo/Redo 상태 저장
            self.save_text_state(device_id)

            # 줄번호 갱신
            self.update_line_numbers(device_id)

            # layout_text에 포커스 주기
            tab_info['layout_text'].focus_set()

            # 임시 파일 삭제
            if os.path.exists(f'temp_layout_{device_id}.xml'):
                os.remove(f'temp_layout_{device_id}.xml')

        except subprocess.CalledProcessError as e:
            messagebox.showerror("오류", f"레이아웃 캡처 실패: {e}")
        except ET.ParseError as e:
            messagebox.showerror("오류", f"XML 파싱 실패: {e}")
        except FileNotFoundError as e:
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다: {e}")
        except Exception as e:
            messagebox.showerror("오류", f"레이아웃 캡처 중 오류 발생: {e}")
            # 임시 파일 정리
            if os.path.exists(f'temp_layout_{device_id}.xml'):
                os.remove(f'temp_layout_{device_id}.xml')

    def save_capture(self, device_id):
        """캡처 저장"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        current_id = tab_info['id_entry'].get()

        if not current_id:
            messagebox.showwarning("경고", "먼저 캡처를 실행하세요.")
            return

        # 현재 프로그램 디렉토리 기준으로 디렉토리 생성
        current_dir = os.path.dirname(os.path.abspath(__file__))
        img_dir = os.path.join(current_dir, "img")
        layout_dir = os.path.join(current_dir, "layout")

        os.makedirs(img_dir, exist_ok=True)
        os.makedirs(layout_dir, exist_ok=True)

        saved_files = []

        try:
            # 이미지 저장 (원본 크기로 저장)
            if tab_info['original_image']:
                img_path = os.path.join(img_dir, f"{current_id}.png")
                tab_info['original_image'].save(img_path)
                saved_files.append(("screen", f"img:{img_path}"))

            # 레이아웃 저장 (현재 편집된 내용 저장)
            layout_content = tab_info['layout_text'].get(1.0, tk.END)
            if layout_content.strip():
                layout_path = os.path.join(layout_dir, f"{current_id}.xml")
                with open(layout_path, 'w', encoding='utf-8') as f:
                    f.write(layout_content)
                saved_files.append(("layout", f"xml:{layout_path}"))

            # 요청사항: 파일 경로를 각각의 라벨에 표시
            for file_type, path_text in saved_files:
                if file_type == "screen":
                    tab_info['screen_path_label'].configure(text=path_text)
                elif file_type == "layout":
                    tab_info['layout_path_label'].configure(text=path_text)

            messagebox.showinfo("성공", f"파일이 저장되었습니다:\n- img/{current_id}.png\n- layout/{current_id}.xml")

        except Exception as e:
            messagebox.showerror("오류", f"파일 저장 실패: {e}")

    def on_canvas_drag(self, event, device_id):
        """캔버스 드래그"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]
        if tab_info['selection_start']:
            if tab_info['selection_rect']:
                tab_info['screen_canvas'].delete(tab_info['selection_rect'])

            canvas = tab_info['screen_canvas']
            x = int(canvas.canvasx(event.x))
            y = int(canvas.canvasy(event.y))

            tab_info['selection_rect'] = tab_info['screen_canvas'].create_rectangle(
                tab_info['selection_start'][0], tab_info['selection_start'][1],
                x, y,
                outline='red', dash=(5, 5), width=2
            )

    def save_selection(self, device_id):
        """선택 영역 저장"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]

        if not tab_info['selection_start'] or not tab_info['selection_end'] or not tab_info['original_image']:
            messagebox.showwarning("경고", "먼저 영역을 선택하세요.")
            return

        # 프로그램 폴더를 기본 저장 위치로 설정
        current_dir = os.path.dirname(os.path.abspath(__file__))

        filename = filedialog.asksaveasfilename(
            initialdir=current_dir,
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
        )

        if filename:
            try:
                # 선택 영역 좌표 정규화 (좌상단, 우하단 순서로)
                x1 = min(tab_info['selection_start'][0], tab_info['selection_end'][0])
                y1 = min(tab_info['selection_start'][1], tab_info['selection_end'][1])
                x2 = max(tab_info['selection_start'][0], tab_info['selection_end'][0])
                y2 = max(tab_info['selection_start'][1], tab_info['selection_end'][1])

                # 선택 영역이 너무 작은 경우 체크
                if (x2 - x1) < 5 or (y2 - y1) < 5:
                    messagebox.showwarning("경고", "선택 영역이 너무 작습니다.")
                    return

                # 축소 비율을 고려하여 원본 이미지 좌표로 변환
                scale_ratio = tab_info['current_scale'] / 100.0
                orig_x1 = max(0, int(x1 / scale_ratio))
                orig_y1 = max(0, int(y1 / scale_ratio))
                orig_x2 = min(tab_info['original_image'].width, int(x2 / scale_ratio))
                orig_y2 = min(tab_info['original_image'].height, int(y2 / scale_ratio))

                # 최종 좌표 검증
                if orig_x1 >= orig_x2 or orig_y1 >= orig_y2:
                    messagebox.showwarning("경고", "잘못된 선택 영역입니다. 다시 선택해주세요.")
                    return

                # 선택 영역 크롭 및 저장
                cropped = tab_info['original_image'].crop((orig_x1, orig_y1, orig_x2, orig_y2))

                # 크롭된 이미지가 비어있는지 확인
                if cropped.width == 0 or cropped.height == 0:
                    messagebox.showwarning("경고", "선택 영역이 비어있습니다.")
                    return

                cropped.save(filename)
                messagebox.showinfo("성공", f"선택 영역이 저장되었습니다:\n{filename}")

                # 선택 영역 초기화
                if tab_info['selection_rect']:
                    tab_info['screen_canvas'].delete(tab_info['selection_rect'])
                    tab_info['selection_rect'] = None
                tab_info['selection_start'] = None
                tab_info['selection_end'] = None

            except Exception as e:
                messagebox.showerror("오류", f"선택 영역 저장 실패: {str(e)}")

    def copy_selection(self, device_id):
        """선택 영역을 클립보드에 복사"""
        if device_id not in self.device_tabs:
            return

        tab_info = self.device_tabs[device_id]

        if not tab_info['selection_start'] or not tab_info['selection_end'] or not tab_info['original_image']:
            messagebox.showwarning("경고", "먼저 영역을 선택하세요.")
            return

        try:
            # save_selection과 동일한 좌표 계산 로직
            x1 = min(tab_info['selection_start'][0], tab_info['selection_end'][0])
            y1 = min(tab_info['selection_start'][1], tab_info['selection_end'][1])
            x2 = max(tab_info['selection_start'][0], tab_info['selection_end'][0])
            y2 = max(tab_info['selection_start'][1], tab_info['selection_end'][1])

            # 선택 영역이 너무 작은 경우 체크
            if (x2 - x1) < 5 or (y2 - y1) < 5:
                messagebox.showwarning("경고", "선택 영역이 너무 작습니다.")
                return

            # 축소 비율을 고려하여 원본 이미지 좌표로 변환
            scale_ratio = tab_info['current_scale'] / 100.0
            orig_x1 = max(0, int(x1 / scale_ratio))
            orig_y1 = max(0, int(y1 / scale_ratio))
            orig_x2 = min(tab_info['original_image'].width, int(x2 / scale_ratio))
            orig_y2 = min(tab_info['original_image'].height, int(y2 / scale_ratio))

            if orig_x1 >= orig_x2 or orig_y1 >= orig_y2:
                messagebox.showwarning("경고", "잘못된 선택 영역입니다. 다시 선택해주세요.")
                return

            # 선택 영역 크롭
            cropped = tab_info['original_image'].crop((orig_x1, orig_y1, orig_x2, orig_y2))

            if cropped.width == 0 or cropped.height == 0:
                messagebox.showwarning("경고", "선택 영역이 비어있습니다.")
                return

            # 클립보드에 복사 (Windows)
            output = BytesIO()
            cropped.save(output, 'BMP')
            data = output.getvalue()[14:]  # BMP 헤더 제거
            output.close()

            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()

            messagebox.showinfo("성공", "선택 영역이 클립보드에 복사되었습니다.")

            # 선택 영역 초기화
            if tab_info['selection_rect']:
                tab_info['screen_canvas'].delete(tab_info['selection_rect'])
                tab_info['selection_rect'] = None
            tab_info['selection_start'] = None
            tab_info['selection_end'] = None

        except ImportError:
            messagebox.showerror("오류", "클립보드 복사를 위해 pywin32 패키지가 필요합니다.\npip install pywin32")
        except Exception as e:
            messagebox.showerror("오류", f"클립보드 복사 실패: {str(e)}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = ScreenLayoutCapture()
    app.run()