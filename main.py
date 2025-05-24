__version__ = '0.1.0'

from math import hypot
from pyautogui import drag, easeOutQuad, leftClick, locateAllOnScreen, locateOnScreen, moveTo, position, screenshot, mouseDown, mouseUp
import fruit_box_bot
import time
from typing import List, Tuple
import win32gui
import win32api
import win32con
import threading
import os
import sys
import tkinter as tk
from tkinter import ttk, scrolledtext
from tkinter import messagebox
import queue
from PIL import Image, ImageTk, ImageGrab
import cv2
import numpy as np
import pyautogui
import json
import logging
from datetime import datetime
import keyboard

# 게임 그리드 크기 (10행 x 17열)
NUM_ROWS = 10
NUM_COLS = 17

# 전역 변수로 프로그램 실행 상태 추가
running = True
game_region = None
is_ready_to_start = False
main_window = None
is_dragging = False  # 드래그 상태 추적을 위한 전역 변수 추가
current_grid = None

def log_message(message, print_to_gui=True):
    """GUI에 로그 메시지를 표시하고 디버그 파일에도 기록 (스레드 안전)"""
    try:
        if main_window and print_to_gui:
            main_window.update_gui(lambda: main_window.log_text.insert(tk.END, message + "\n"))
            main_window.update_gui(lambda: main_window.log_text.see(tk.END))
        
        # 디버그 파일에 로그 기록
        try:
            with open('debug.txt', 'a', encoding='utf-8') as f:
                f.write(message + "\n")
        except Exception as file_error:
            print(f"로그 파일 기록 중 오류: {str(file_error)}")
    except Exception as log_error:
        print(f"로그 메시지 처리 중 오류: {str(log_error)}")

def print_apple_grid(grid):
    """사과 번호 매핑 테이블을 출력하고 파일에 기록"""
    grid_str = "\n=== 사과 번호 매핑 테이블 ===\n"
    for row in grid:
        row_str = " ".join(str(cell).rjust(2) if cell != 0 else "·".rjust(2) for cell in row)
        grid_str += f"{row_str}\n"
    
    # GUI와 파일 모두에 출력
    log_message(grid_str, print_to_gui=True)

def check_keys():
    """전역 키보드 이벤트를 체크하는 스레드 함수"""
    global running, is_ready_to_start
    while True:
        try:
            # F2 키 체크 (VK_F2 = 0x71)
            if win32api.GetAsyncKeyState(0x71) & 0x8000:
                if game_region is None:
                    if main_window:
                        main_window.update_gui(lambda: messagebox.showwarning("알림", "게임 영역을 먼저 선택해주세요!"))
                else:
                    running = True
                    is_ready_to_start = True
                    if main_window:
                        main_window.update_gui(lambda: main_window.log_message("\n게임을 시작!(알고리즘 6~10초 소요)"))
                time.sleep(0.5)  # 키 입력 중복 방지
            
            # ESC 키 체크 (VK_ESCAPE = 0x1B)
            if win32api.GetAsyncKeyState(0x1B) & 0x8000:
                if main_window:
                    main_window.update_gui(lambda: main_window.stop_mouse())
                    mouseUp(button='left')  # 강제로 마우스 업
                    time.sleep(0.1)
                    running = False  # running 상태를 즉시 False로
                time.sleep(0.5)  # 키 입력 중복 방지

            # F4 키 체크 (VK_F4 = 0x73)
            if win32api.GetAsyncKeyState(0x73) & 0x8000:
                if main_window:
                    main_window.update_gui(lambda: main_window.stop_mouse())
                    mouseUp(button='left')  # 강제로 마우스 업
                time.sleep(0.5)  # 키 입력 중복 방지

            # 넘버패드 0 키 체크 (VK_NUMPAD0 = 0x60)
            if win32api.GetAsyncKeyState(0x60) & 0x8000:
                if main_window:
                    main_window.update_gui(lambda: main_window.test_array_display())
                time.sleep(0.5)  # 키 입력 중복 방지
            
            time.sleep(0.1)
        except Exception as e:
            if main_window:
                main_window.update_gui(lambda: main_window.log_message(f"\n키 체크 중 오류 발생: {e}"))
            time.sleep(1)

def merge_nearby_positions(positions, threshold=5):
    """인접한 위치들을 하나로 병합"""
    if not positions:
        return []
    
    merged = []
    current_group = [positions[0]]
    
    for pos in positions[1:]:
        last_pos = current_group[-1]
        if abs(pos[0] - last_pos[0]) <= threshold and abs(pos[1] - last_pos[1]) <= threshold:
            current_group.append(pos)
        else:
            if current_group:
                avg_x = sum(p[0] for p in current_group) // len(current_group)
                avg_y = sum(p[1] for p in current_group) // len(current_group)
                merged.append((avg_x, avg_y, current_group[0][2], current_group[0][3]))
            current_group = [pos]
    
    if current_group:
        avg_x = sum(p[0] for p in current_group) // len(current_group)
        avg_y = sum(p[1] for p in current_group) // len(current_group)
        merged.append((avg_x, avg_y, current_group[0][2], current_group[0][3]))
    
    return merged

class TransparentWindow:
    def __init__(self):
        self.root = tk.Toplevel()
        self.root.title("게임 영역 선택")
        
        # 창을 완전 투명하게 설정
        self.root.attributes('-alpha', 1.0)
        self.root.attributes('-topmost', True)
        self.root.configure(bg='white')
        self.root.wm_attributes('-transparentcolor', 'white')
        
        # 창 크기 조절 가능하게 설정
        self.root.resizable(True, True)
        
        # 최소 크기 설정
        self.root.minsize(200, 200)
        
        # 버튼 프레임 생성 (버튼용)
        self.button_frame = ttk.Frame(self.root)
        self.button_frame.pack(fill='x', side='bottom')
        
        # 확인 버튼
        self.confirm_button = ttk.Button(
            self.button_frame, 
            text="영역 선택 완료 (Space)", 
            command=self.confirm_selection
        )
        self.confirm_button.pack(pady=5)
        
        # 버튼 프레임의 높이를 구하기 위해 업데이트
        self.root.update()
        button_height = self.button_frame.winfo_height()
        
        # 시작 크기와 위치 설정 (테스트된 크기로 설정)
        if hasattr(self.root.master, 'tested_width'):
            width = self.root.master.tested_width
            height = self.root.master.tested_height
        else:
            width = 1115  # 기본값
            height = 650  # 기본값
            
        # 버튼 영역 높이를 고려하여 전체 창 크기 조정
        total_height = height + button_height
            
        # 화면 중앙에 위치하도록 계산
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - total_height) // 2
        
        self.root.geometry(f"{width}x{total_height}+{x}+{y}")
        
        # 캔버스 생성 (버튼 프레임 위에)
        self.canvas = tk.Canvas(self.root, bg='white', highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)
        
        # 드래그 시작 좌표
        self.start_x = None
        self.start_y = None
        self.dragging = False
        
        # 마우스 이벤트 바인딩
        self.canvas.bind('<Button-1>', self.start_drag)
        self.canvas.bind('<B1-Motion>', self.update_drag)
        self.canvas.bind('<ButtonRelease-1>', self.end_drag)
        
        # 메인 창 참조 저장
        self.main_window = self.root.master
        
        # 테두리 그리기 (빨간색, 두께 2)
        def update_border(event=None):
            self.canvas.delete('border')
            width = self.canvas.winfo_width()
            height = self.canvas.winfo_height()
            self.canvas.create_rectangle(2, 2, width-2, height-2, 
                                      outline='red', width=2, 
                                      tags='border')
        
        self.canvas.bind('<Configure>', update_border)
        
        # Space 키 바인딩
        self.root.bind('<space>', lambda e: self.confirm_selection())
        
        self.result = None
        
        # 초기 테두리 그리기
        self.root.update()
        update_border()
    
    def start_drag(self, event):
        """드래그 시작"""
        self.start_x = event.x
        self.start_y = event.y
        self.dragging = True
    
    def update_drag(self, event):
        """드래그 중 업데이트"""
        if self.dragging:
            # 드래그 영역 표시
            self.canvas.delete('drag_area')
            self.canvas.create_rectangle(
                self.start_x, self.start_y, event.x, event.y,
                outline='blue', width=2,
                tags='drag_area'
            )
            
            # 드래그 영역의 크기 계산
            width = abs(event.x - self.start_x)
            height = abs(event.y - self.start_y)
            
            # 예상 점수 계산
            cells_width = int(width / (self.canvas.winfo_width() / NUM_COLS))
            cells_height = int(height / (self.canvas.winfo_height() / NUM_ROWS))
            estimated_score = cells_width * cells_height
            
            # 메인 창의 상태 표시 업데이트
            if hasattr(main_window, 'status_var'):
                main_window.update_gui(
                    lambda: main_window.status_var.set(
                        f"예상 영역: {cells_width}x{cells_height} 셀 (점수: {estimated_score})"
                    )
                )
    
    def end_drag(self, event):
        """드래그 종료"""
        self.dragging = False
    
    def confirm_selection(self):
        # 창의 위치와 크기 저장 (타이틀바 제외, 테두리 포함)
        hwnd = win32gui.FindWindow(None, "게임 영역 선택")
        
        # 전체 창의 크기와 위치
        window_rect = win32gui.GetWindowRect(hwnd)
        
        # 타이틀바 높이 계산 (기본 캡션 높이만 사용)
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if style & win32con.WS_CAPTION:
            title_height = win32api.GetSystemMetrics(win32con.SM_CYCAPTION)
        else:
            title_height = 0
        
        # 버튼 프레임의 높이 계산
        button_height = self.button_frame.winfo_height()
            
        # 타이틀바만 제외하고 버튼 영역 위까지만 계산
        x = window_rect[0]
        y = window_rect[1] + title_height
        width = window_rect[2] - window_rect[0]
        height = window_rect[3] - window_rect[1] - title_height - button_height
        
        self.result = (x, y, width, height)
        self.root.quit()
    
    def get_region(self):
        self.root.mainloop()
        self.root.destroy()
        return self.result

def create_scaled_templates():
    """기준 크기에 따른 스케일링된 템플릿 이미지 생성"""
    try:
        # 스케일링된 이미지를 저장할 디렉토리 생성
        os.makedirs('scaled_images/small', exist_ok=True)
        os.makedirs('scaled_images/medium', exist_ok=True)
        os.makedirs('scaled_images/large', exist_ok=True)

        # 기준 크기 (medium)
        REFERENCE_WIDTH = 1133
        REFERENCE_HEIGHT = 672

        # 스케일 비율 설정
        SMALL_SCALE = 0.7  # 30% 작게
        LARGE_SCALE = 1.3  # 30% 크게

        log_message("\n템플릿 이미지 스케일링 중...", print_to_gui=True)

        for digit in range(1, 10):
            try:
                image_path = resource_path(f'images/apple{digit}.png')
                template = Image.open(image_path)

                # 작은 크기 템플릿 생성
                small_size = (
                    int(template.width * SMALL_SCALE),
                    int(template.height * SMALL_SCALE)
                )
                small_template = template.resize(small_size, Image.Resampling.LANCZOS)
                small_template.save(f'scaled_images/small/apple{digit}.png')

                # 원본 크기 템플릿 복사
                template.save(f'scaled_images/medium/apple{digit}.png')

                # 큰 크기 템플릿 생성
                large_size = (
                    int(template.width * LARGE_SCALE),
                    int(template.height * LARGE_SCALE)
                )
                large_template = template.resize(large_size, Image.Resampling.LANCZOS)
                large_template.save(f'scaled_images/large/apple{digit}.png')

            except Exception as e:
                log_message(f"사과 {digit} 템플릿 스케일링 중 오류: {e}", print_to_gui=True)

        log_message("템플릿 이미지 스케일링 완료!", print_to_gui=True)
        return True

    except Exception as e:
        log_message(f"템플릿 생성 중 오류 발생: {e}", print_to_gui=True)
        return False

def get_template_size_category(width, height):
    """게임 영역 크기에 따른 템플릿 크기 카테고리 반환"""
    # 기준 크기
    REFERENCE_WIDTH = 1133
    REFERENCE_HEIGHT = 672

    # 영역 크기와 기준 크기의 비율 계산
    width_ratio = width / REFERENCE_WIDTH
    height_ratio = height / REFERENCE_HEIGHT
    avg_ratio = (width_ratio + height_ratio) / 2

    # 카테고리 결정 (기준 수정)
    if avg_ratio < 0.95:  # 기준보다 5% 이상 작으면
        return 'small'
    elif avg_ratio > 1.05:  # 기준보다 5% 이상 크면
        return 'large'
    else:  # 그 외에는 중간 크기
        return 'medium'

def load_scaled_template(digit, scale, tested_width, tested_height):
    """스케일에 맞춰 템플릿 이미지를 리사이즈"""
    img = Image.open(resource_path(f'images/apple{digit}.png'))
    # 원본 템플릿 크기
    original_width, original_height = img.size
    
    # 현재 게임 영역과 기준 크기의 비율 계산
    width_scale = game_region[2] / tested_width
    height_scale = game_region[3] / tested_height
    
    # 가로, 세로 비율 중 더 작은 값을 사용 (비율 유지)
    final_scale = min(width_scale, height_scale)
    
    # 새로운 크기 계산
    new_width = int(original_width * final_scale)
    new_height = int(original_height * final_scale)
    
    return img.resize((new_width, new_height), Image.Resampling.LANCZOS)

class MainWindow:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("과일 박스 게임 봇")
        self.window.geometry("800x600")
        
        # 로깅 설정
        logging.basicConfig(filename='debug.txt', level=logging.DEBUG,
                          format='%(asctime)s - %(levelname)s - %(message)s')
        
        # 초기 속성 설정
        self.mouse_speed = 1.0
        self.reference_width = None
        self.reference_height = None
        self.update_queue = queue.Queue()
        self.message_queue = queue.Queue()
        
        self.setup_ui()
        self.running = False
        self.paused = False
        self.current_grid = None
        
        # 설정 파일 경로
        self.config_file = 'config.json'
        
    def setup_ui(self):
        # 스타일 설정
        style = ttk.Style()
        style.configure("TButton", padding=5)
        
        # 메인 프레임
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # 상단 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(0, 5))
        
        # 영역 선택 버튼
        self.select_area_btn = ttk.Button(
            button_frame, 
            text="게임 영역 선택", 
            command=self.start_area_selection
        )
        self.select_area_btn.pack(side='left', padx=5)
        
        # 예시 보기 버튼
        self.example_btn = ttk.Button(
            button_frame,
            text="영역 선택 예시 보기",
            command=self.show_example
        )
        self.example_btn.pack(side='left', padx=5)
        
        # 시작 버튼
        self.start_btn = ttk.Button(
            button_frame,
            text="게임 시작 (F2)",
            command=self.start_game
        )
        self.start_btn.pack(side='left', padx=5)
        self.start_btn['state'] = 'disabled'
        
        # 마우스 멈춤 버튼
        self.stop_btn = ttk.Button(
            button_frame,
            text="리셋/멈춤 (F4)",
            command=self.stop_mouse
        )
        self.stop_btn.pack(side='left', padx=5)

        # 마우스 속도 조절 프레임
        speed_frame = ttk.Frame(main_frame)
        speed_frame.pack(fill='x', pady=(0, 5))
        
        # 속도 조절 레이블
        speed_label_text = ttk.Label(
            speed_frame,
            text="마우스 속도 조절:",
            width=15
        )
        speed_label_text.pack(side='left', padx=5)
        
        # 속도 감소 버튼
        self.speed_down_btn = ttk.Button(
            speed_frame,
            text="속도 감소 (-)",
            command=self.decrease_speed,
            width=12
        )
        self.speed_down_btn.pack(side='left', padx=2)
        
        # 현재 속도 표시 레이블
        self.speed_label = ttk.Label(
            speed_frame,
            text="속도: 1.0x",
            width=10
        )
        self.speed_label.pack(side='left', padx=5)
        
        # 속도 증가 버튼
        self.speed_up_btn = ttk.Button(
            speed_frame,
            text="속도 증가 (+)",
            command=self.increase_speed,
            width=12
        )
        self.speed_up_btn.pack(side='left', padx=2)

        # 글씨 크기 조절 및 테스트 프레임
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill='x', pady=(0, 5))

        # 글씨 크기 조절 레이블
        font_label = ttk.Label(
            control_frame,
            text="글씨 크기:",
            width=10
        )
        font_label.pack(side='left', padx=5)

        # 글씨 크기 감소 버튼
        self.font_down_btn = ttk.Button(
            control_frame,
            text="작게 (-)",
            command=self.decrease_font_size,
            width=8
        )
        self.font_down_btn.pack(side='left', padx=2)

        # 현재 글씨 크기 표시 레이블
        self.font_size = 10  # 기본 글씨 크기
        self.font_label = ttk.Label(
            control_frame,
            text=f"크기: {self.font_size}",
            width=8
        )
        self.font_label.pack(side='left', padx=5)

        # 글씨 크기 증가 버튼
        self.font_up_btn = ttk.Button(
            control_frame,
            text="크게 (+)",
            command=self.increase_font_size,
            width=8
        )
        self.font_up_btn.pack(side='left', padx=2)

        # 배열 테스트 버튼
        self.test_array_btn = ttk.Button(
            control_frame,
            text="배열 테스트 (Num 0)",
            command=self.test_array_display,
            width=18
        )
        self.test_array_btn.pack(side='left', padx=10)
        
        # 로그 창
        self.log_text = scrolledtext.ScrolledText(main_frame, height=15)
        self.log_text.pack(fill='both', expand=True)
        
        # 상태 표시줄
        self.status_var = tk.StringVar()
        self.status_var.set("게임 영역을 선택해주세요")
        status_label = ttk.Label(
            main_frame, 
            textvariable=self.status_var,
            relief='sunken',
            padding=5
        )
        status_label.pack(fill='x', pady=(10, 0))
        
        # 제작자 정보
        creator_label = ttk.Label(
            main_frame,
            text="제작: 거누@https://github.com/cgw0331",
            foreground='gray'
        )
        creator_label.pack(pady=(5, 0))
        
        # 초기 메시지 표시
        self.show_initial_message()
        
        # 스레드 안전한 GUI 업데이트를 위한 큐
        self.check_queue()
        self.total_apples = 0  # 인식된 총 사과 수를 저장할 변수 추가

    def log_message(self, message, print_to_gui=True):
        """GUI에 로그 메시지를 표시하고 디버그 파일에도 기록"""
        try:
            if print_to_gui and hasattr(self, 'log_text'):
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
            
            # 디버그 파일에도 기록
            try:
                with open('debug.txt', 'a', encoding='utf-8') as f:
                    f.write(message + "\n")
            except Exception as e:
                print(f"로그 파일 기록 중 오류: {e}")
        except Exception as e:
            print(f"로그 메시지 처리 중 오류: {e}")

    def load_templates(self):
        """숫자 이미지 템플릿을 로드합니다."""
        self.templates = {}
        template_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'images')
        for i in range(10):
            template_path = os.path.join(template_dir, f'{i}.png')
            if os.path.exists(template_path):
                self.templates[i] = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)

    def recognize_digit(self, image):
        """이미지에서 숫자를 인식합니다."""
        # 그레이스케일로 변환
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # 결과 저장을 위한 변수
        best_match = None
        best_score = float('-inf')
        recognized_digit = None
        
        # 각 템플릿과 매칭
        for digit, template in self.templates.items():
            # 템플릿 매칭
            result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
            score = result.max()
            
            # 가장 높은 점수를 가진 매칭 저장
            if score > best_score:
                best_score = score
                recognized_digit = digit
        
        # 신뢰도가 0.8 이상일 때만 인식 결과 반환
        if best_score >= 0.8:
            return recognized_digit
        return None

    def show_initial_message(self):
        """초기 메시지를 표시"""
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "=== Fruit Box 게임 봇 설명서 ===\n")
        self.log_text.insert(tk.END, "1. '게임 영역 선택' 버튼을 클릭하여 게임 영역을 선택해야함(예시 참고)\n")
        self.log_text.insert(tk.END, "2. 영역 선택 후 자동으로 사과 인식 테스트를 진행합니다.\n")
        self.log_text.insert(tk.END, "3. 테스트가 성공하면 F2를 눌러 게임을 시작할 수 있습니다.\n")
        self.log_text.insert(tk.END, "4. 게임 도중 ESC 또는 F4를 누루면 초기화됨(좀 잘씹힘, 될때까지 연타).\n")
        self.log_text.insert(tk.END, "\n[주의사항]\n")
        self.log_text.insert(tk.END, "- 마우스 속도는 +/- 버튼으로 조절할 수 있습니다.\n")
        self.log_text.insert(tk.END, "- 해상도 상관없이 동작하도록 코드작성하긴 했는데.. 버그 있을 수 있음)\n")
        self.log_text.insert(tk.END, "- 게임 영역 선택 시, 사과를 콤팩트하게 설정해야함.\n")
        self.log_text.insert(tk.END, "- 사과는 총 170개가 인식되어야 함, 안되면 브라우저 크기 좀 조절하고 다시해보셈.\n")
        self.log_text.insert(tk.END, "- 게임 영역을 선택한 후에는 게임 화면의 위치를 바꾸면 안됨.\n")
        self.log_text.insert(tk.END, "- 게임 화면의 위치가 바뀌었다면 게임 영역을 다시 선택하여야함(절대 좌표 기준).\n")
        self.log_text.insert(tk.END, "- 게임 실행 중에는 마우스 가만히 냅두셈..\n")
        self.log_text.see(tk.END)

    def test_apple_recognition(self):
        """사과 숫자 인식 테스트"""
        global game_region
        
        if game_region is None:
            messagebox.showwarning("알림", "게임 영역을 먼저 선택해주세요!")
            return None
        
        try:
            # 기준 크기 설정 (처음 테스트 성공했던 크기)
            tested_width = 1115
            tested_height = 650
            
            # 현재 게임 영역의 스케일 계산
            scale = min(game_region[2]/tested_width, game_region[3]/tested_height)
            
            # 그리드 초기화
            grid = [[0 for _ in range(NUM_COLS)] for _ in range(NUM_ROWS)]
            
            # 각 숫자에 대해 템플릿 매칭 수행
            for digit in range(1, 10):
                try:
                    # 스케일된 템플릿 로드
                    tmpl = load_scaled_template(digit, scale, tested_width, tested_height)
                    
                    # pyautogui 템플릿 매칭
                    matches = list(pyautogui.locateAllOnScreen(
                        tmpl, 
                        confidence=0.93,
                        grayscale=True,
                        region=game_region
                    ))
                    
                    # 인접한 매칭 결과 병합
                    matches = merge_nearby_positions(matches, threshold=4)
                    
                    # 매칭된 위치를 그리드에 매핑
                    for match in matches:
                        col = int((match[0] - game_region[0]) / (game_region[2] / NUM_COLS))
                        row = int((match[1] - game_region[1]) / (game_region[3] / NUM_ROWS))
                        if 0 <= row < NUM_ROWS and 0 <= col < NUM_COLS and grid[row][col] == 0:
                            grid[row][col] = digit
                            
                except Exception as e:
                    self.log_message(f"사과 {digit} 인식 중 오류: {e}")
                    continue
            
            # 그리드 출력
            print_apple_grid(grid)
            
            # 인식된 숫자 개수 확인
            total_numbers = sum(1 for row in grid for cell in row if cell > 0)
            self.log_message(f"\n총 {total_numbers}개의 숫자를 인식했습니다.")
            
            if total_numbers != 170:
                self.log_message(f"\n경고: 숫자 개수가 맞지 않습니다! (현재: {total_numbers}, 필요: 170)")
                self.log_message("브라우저 확대/축소나 게임 영역 크기를 조절해보세요.")
            
            # 현재 그리드 상태 업데이트
            self.current_grid = grid
            return grid
            
        except Exception as e:
            self.log_message(f"\n숫자 인식 테스트 중 오류 발생: {e}")
            return None

    def start_area_selection(self):
        """게임 영역 선택 시작"""
        global game_region
        window = TransparentWindow()
        game_region = window.get_region()
        
        if game_region:
            # 영역 선택 결과만 표시
            self.log_message(f"\n선택된 게임 영역:")
            self.log_message(f"위치: ({game_region[0]}, {game_region[1]})")
            self.log_message(f"크기: {game_region[2]}x{game_region[3]}")
            
            # F2 키를 눌러 시작하도록 안내
            self.start_btn['state'] = 'normal'
            self.status_var.set("F2 키를 눌러 게임을 시작할 수 있습니다")
        else:
            self.log_message("게임 영역이 선택되지 않았습니다.")

    def reset_program(self):
        """프로그램을 초기 상태로 되돌림 (게임 영역 좌표 제외)"""
        global running, is_ready_to_start
        
        # 실행 상태 초기화
        running = False
        is_ready_to_start = False
        
        # 마우스 상태 초기화
        mouseUp(button='left')
        time.sleep(0.2)  # 마우스 동작이 완전히 중지될 때까지 대기
        
        # 새로운 시작 준비
        running = True
        
        # 버튼 상태 초기화
        if game_region is not None:
            self.start_btn['state'] = 'normal'
            self.status_var.set("F2 키를 눌러 게임을 다시 시작할 수 있습니다")
        else:
            self.start_btn['state'] = 'disabled'
            self.status_var.set("게임 영역을 선택해주세요")
        
        # 로그창 초기화
        self.show_initial_message()
        
        # 게임 영역 정보 표시 (있는 경우)
        if game_region:
            self.log_message(f"\n현재 선택된 게임 영역:")
            self.log_message(f"위치: ({game_region[0]}, {game_region[1]})")
            self.log_message(f"크기: {game_region[2]}x{game_region[3]}")
            self.log_message("\nF2 키를 눌러 게임을 시작할 수 있습니다")

    def stop_mouse(self):
        """마우스 멈춤 버튼 클릭 핸들러"""
        global is_dragging
        is_dragging = False  # 드래그 상태 즉시 해제
        pyautogui.mouseUp(button='left')  # mouse.release() 대신 pyautogui 사용
        self.reset_program()  # reset_program 함수 호출

    def start_game(self):
        """게임 시작 버튼 클릭 핸들러"""
        global running, is_ready_to_start
        if game_region is None:
            messagebox.showwarning("알림", "게임 영역을 먼저 선택해주세요!")
            return
        
        # 시작 전에 프로그램 상태 초기화
        self.reset_program()
        
        # 게임 시작
        is_ready_to_start = True
        self.log_message("\n게임을 시작합니다! (알고리즘 6~10초 소요)")

    def check_queue(self):
        """GUI 업데이트 큐를 확인하고 처리"""
        try:
            while not self.update_queue.empty():
                try:
                    func = self.update_queue.get_nowait()
                    if func:
                        func()
                except Exception as e:
                    print(f"함수 실행 중 오류: {str(e)}")
        except Exception as e:
            print(f"큐 처리 중 오류: {str(e)}")
        finally:
            # 계속해서 큐를 확인
            self.window.after(100, self.check_queue)

    def update_gui(self, func):
        """스레드 안전한 GUI 업데이트"""
        try:
            if running:
                self.update_queue.put(func)
        except Exception as update_error:
            print(f"GUI 업데이트 중 오류: {str(update_error)}")

    def show_example(self):
        """영역 선택 예시 이미지를 보여줌"""
        example_window = tk.Toplevel(self.window)
        example_window.title("영역 선택 예시")
        example_window.attributes('-topmost', True)
        
        try:
            image_path = resource_path('images/region_example.png')
            image = tk.PhotoImage(file=image_path)
            label = ttk.Label(example_window, image=image)
            label.image = image  # 참조 유지
            label.pack(padx=10, pady=10)
            
            # 설명 텍스트
            ttk.Label(
                example_window,
                text="빨간색 테두리를 게임 영역에 맞추세요",
                wraplength=300,
                justify='center'
            ).pack(pady=(0, 10))
            
        except Exception as e:
            ttk.Label(
                example_window,
                text="예시 이미지를 불러올 수 없습니다.",
                wraplength=300,
                justify='center'
            ).pack(padx=20, pady=20)
            self.log_message(f"예시 이미지 로드 오류: {e}")

    def __del__(self):
        """소멸자"""
        pass

    def increase_speed(self):
        """마우스 속도 증가"""
        self.mouse_speed = min(2.0, self.mouse_speed + 0.1)
        self.update_speed_label()
        self.log_message(f"\n마우스 속도가 {self.mouse_speed:.1f}배로 증가했습니다.")
    
    def decrease_speed(self):
        """마우스 속도 감소"""
        self.mouse_speed = max(0.5, self.mouse_speed - 0.1)
        self.update_speed_label()
        self.log_message(f"\n마우스 속도가 {self.mouse_speed:.1f}배로 감소했습니다.")
    
    def update_speed_label(self):
        """속도 레이블 업데이트"""
        self.speed_label.config(text=f"속도: {self.mouse_speed:.1f}x")

    def increase_font_size(self):
        """글씨 크기 증가"""
        self.font_size = min(20, self.font_size + 1)  # 최대 20
        self.update_font_size()
        self.log_message(f"\n글씨 크기가 {self.font_size}로 증가했습니다.")
    
    def decrease_font_size(self):
        """글씨 크기 감소"""
        self.font_size = max(8, self.font_size - 1)  # 최소 8
        self.update_font_size()
        self.log_message(f"\n글씨 크기가 {self.font_size}로 감소했습니다.")
    
    def update_font_size(self):
        """글씨 크기 업데이트"""
        self.font_label.config(text=f"크기: {self.font_size}")
        self.log_text.configure(font=('TkDefaultFont', self.font_size))
    
    def test_array_display(self):
        """배열 테스트 실행"""
        if game_region is None:
            messagebox.showwarning("알림", "게임 영역을 먼저 선택해주세요!")
            return
        
        try:
            self.log_message("\n=== 배열 테스트 시작 ===")
            grid = self.test_apple_recognition()
            if grid is not None:
                self.log_message("\n테스트가 완료되었습니다!")
            else:
                self.log_message("\n테스트 실패: 배열을 인식할 수 없습니다.")
        except Exception as e:
            self.log_message(f"\n테스트 중 오류 발생: {e}")

    def preprocess_grid(self, grid):
        """그리드를 전처리하여 더 나은 결과를 얻도록 함"""
        rows = len(grid)
        cols = len(grid[0])
        processed = [[0 for _ in range(cols)] for _ in range(rows)]
        
        # 1단계: 기존 그리드 복사
        for i in range(rows):
            for j in range(cols):
                processed[i][j] = grid[i][j]
        
        # 2단계: 빈 칸 주변 탐색하여 채우기
        for i in range(rows):
            for j in range(cols):
                if processed[i][j] == 0:
                    # 주변 8방향 탐색
                    neighbors = []
                    for di in [-1, 0, 1]:
                        for dj in [-1, 0, 1]:
                            if di == 0 and dj == 0:
                                continue
                            ni, nj = i + di, j + dj
                            if 0 <= ni < rows and 0 <= nj < cols and processed[ni][nj] > 0:
                                neighbors.append(processed[ni][nj])
                    
                    # 주변에 숫자가 있으면 가장 많이 나온 숫자로 채움
                    if neighbors:
                        from collections import Counter
                        counter = Counter(neighbors)
                        processed[i][j] = counter.most_common(1)[0][0]
        
        return processed

    def play_game(self):
        """게임 시작"""
        print("play_game 메서드 시작")  # 디버그 로그
        
        try:
            # 그리드 전처리
            print("그리드 전처리 시작")  # 디버그 로그
            processed_grid = self.preprocess_grid(self.current_grid)
            print("그리드 전처리 완료")  # 디버그 로그
            
            # 전처리된 그리드로 전략 계산
            self.update_queue.put(lambda: self.log_message("\n전처리된 그리드로 전략 계산 중..."))
            print("전략 계산 시작")  # 디버그 로그
            strategy = fruit_box_bot.find_strategy(processed_grid)
            print(f"전략 계산 완료: {strategy.score}점")  # 디버그 로그
            
            # GUI 업데이트는 큐를 통해 수행
            self.update_queue.put(lambda: self.log_message(f"\n전략 계산 완료! 예상 점수: {strategy.score}점"))
            self.update_queue.put(lambda: self.log_message("\n게임 플레이를 시작합니다..."))
            
            # 전략 실행
            print("전략 실행 시작")  # 디버그 로그
            self.execute_strategy(strategy)
            print("전략 실행 완료")  # 디버그 로그
            
        except Exception as e:
            error_msg = str(e)
            print(f"play_game 오류: {error_msg}")  # 디버그 로그
            self.update_queue.put(lambda: self.log_message(f"\n전략 계산 중 오류 발생: {error_msg}"))
            self.update_queue.put(lambda: messagebox.showerror("오류", f"전략 계산 중 오류가 발생했습니다: {error_msg}"))

    def execute_strategy(self, strategy):
        """전략을 실행하는 함수"""
        print("execute_strategy 메서드 시작")  # 디버그 로그
        global running, is_dragging
        
        if not running or not game_region:
            print("실행 조건 미충족")  # 디버그 로그
            return
        
        left, top = game_region[0], game_region[1]
        cell_width, cell_height = calculate_cell_size(game_region)
        self.update_queue.put(lambda: self.log_message(f"\n알고리즘이 계산한 예상 점수: {strategy.score}점"))
        self.update_queue.put(lambda: self.log_message("\n게임 플레이를 시작합니다..."))
        print(f"박스 개수: {len(strategy.boxes)}")  # 디버그 로그

        for box in strategy.boxes:
            if not running:
                mouseUp(button='left')
                is_dragging = False
                break
            
            try:
                # 시작 위치 계산
                start_x = int(left + (box.x * cell_width))
                start_y = int(top + (box.y * cell_height))
                
                # 드래그할 거리 계산
                drag_width = int(box.width * cell_width)
                drag_height = int(box.height * cell_height)
                
                # 드래그 방향에 따른 보정값 계산 (절대값으로 보정)
                CORRECTION_PIXELS = 10
                if drag_width > 0 and drag_height > 0:  # 오른쪽 아래 방향
                    drag_width += CORRECTION_PIXELS
                    drag_height += CORRECTION_PIXELS
                elif drag_width > 0 and drag_height < 0:  # 오른쪽 위 방향
                    drag_width += CORRECTION_PIXELS
                    drag_height -= CORRECTION_PIXELS
                
                # 현재 마우스 위치 확인
                current_x, current_y = position()
                
                if not running:
                    mouseUp(button='left')
                    is_dragging = False
                    break
                
                # 시작 위치로 빠르게 이동
                moveTo(start_x, start_y)
                
                # 이동 거리에 따른 대기 시간 계산
                move_distance = hypot(current_x - start_x, current_y - start_y)
                base_move_wait = 0.02
                move_wait_time = min(0.05, max(0.02, base_move_wait * (move_distance / 200))) / self.mouse_speed
                
                # 이동 후 대기
                start_time = time.time()
                while time.time() - start_time < move_wait_time:
                    if not running:
                        mouseUp(button='left')
                        is_dragging = False
                        break
                    time.sleep(0.01)
                
                if not running:
                    mouseUp(button='left')
                    is_dragging = False
                    break
                
                # 드래그 거리 계산
                drag_distance = hypot(drag_width, drag_height)
                
                # 드래그 거리에 따른 duration 계산
                base_duration = 0.05
                drag_duration = min(0.15, max(base_duration, base_duration * (drag_distance / 100))) / self.mouse_speed
                
                # 목표 지점 계산
                end_x = start_x + drag_width
                end_y = start_y + drag_height
                
                if not running:
                    mouseUp(button='left')
                    is_dragging = False
                    break
                
                # 드래그 시작
                mouseDown(button='left')
                is_dragging = True
                
                # 클릭 후 대기
                click_wait = 0.02
                start_time = time.time()
                while time.time() - start_time < click_wait:
                    if not running:
                        mouseUp(button='left')
                        is_dragging = False
                        break
                    time.sleep(0.01)
                
                if not running:
                    mouseUp(button='left')
                    is_dragging = False
                    break
                
                # 목표 지점으로 이동
                moveTo(end_x, end_y, duration=drag_duration, tween=easeOutQuad)
                
                if not running:
                    mouseUp(button='left')
                    is_dragging = False
                    break
                
                # 목표 지점에서 마우스를 1픽셀 우측으로 살짝 이동
                moveTo(end_x + 1, end_y, duration=0.005)
                
                # 목표 지점에서 대기
                base_hold_time = 0.05
                distance_factor = drag_distance / 300
                hold_time = min(0.1, base_hold_time + distance_factor)
                
                start_time = time.time()
                while time.time() - start_time < hold_time:
                    if not running:
                        mouseUp(button='left')
                        is_dragging = False
                        break
                    time.sleep(0.01)
                
                # 드래그 종료
                mouseUp(button='left')
                is_dragging = False
                
                if not running:
                    break
                
                # 다음 동작 전 대기
                wait_time = 0.1
                
                start_time = time.time()
                while time.time() - start_time < wait_time:
                    if not running:
                        break
                    time.sleep(min(0.01, wait_time/2))
                
            except Exception as e:
                self.log_message(f"\n드래그 중 오류 발생: {e}")
                mouseUp(button='left')
                is_dragging = False

def main():
    print("프로그램 시작...")
    try:
        print("메인 윈도우 생성 시도...")
        global main_window
        main_window = MainWindow()
        print("메인 윈도우 생성 완료")
        
        print("키 체크 스레드 시작...")
        # 키 체크 스레드 시작
        key_thread = threading.Thread(target=check_keys, daemon=True)
        key_thread.start()
        print("키 체크 스레드 시작 완료")
        
        print("게임 로직 스레드 시작...")
        # 게임 로직 스레드 시작
        game_thread = threading.Thread(target=game_loop, daemon=True)
        game_thread.start()
        print("게임 로직 스레드 시작 완료")
        
        print("GUI 메인 루프 시작...")
        # GUI 메인 루프 시작
        main_window.window.mainloop()
    except Exception as e:
        print(f"프로그램 실행 중 오류 발생: {e}")
        if main_window:
            main_window.update_gui(lambda: main_window.log_message(f"\n프로그램 실행 중 오류 발생: {e}"))

def game_loop():
    """게임 실행 로직"""
    global running, game_region, is_ready_to_start, main_window

    while True:
        try:
            if is_ready_to_start and game_region and running and main_window:
                try:
                    # 숫자 인식 시작
                    main_window.current_grid = main_window.test_apple_recognition()
                    
                    if main_window.current_grid is None:
                        error_msg = "숫자 인식에 실패했습니다. 다시 시도해주세요."
                        main_window.update_queue.put(lambda: main_window.log_message(f"\n{error_msg}"))
                        is_ready_to_start = False
                        continue
                    
                    # 게임 실행
                    main_window.play_game()
                    
                except Exception as e:
                    error_msg = str(e)
                    print(f"게임 실행 중 오류: {error_msg}")
                    if main_window:
                        main_window.update_queue.put(lambda: main_window.log_message(f"\n게임 실행 중 오류 발생: {error_msg}"))
                finally:
                    is_ready_to_start = False
                    if main_window:
                        main_window.update_queue.put(lambda: main_window.start_btn.configure(state='normal'))
                        main_window.update_queue.put(lambda: main_window.status_var.set("F2 키를 눌러 게임을 다시 시작할 수 있습니다"))
            time.sleep(0.1)
        except Exception as e:
            error_msg = str(e)
            print(f"게임 루프 오류: {error_msg}")
            if main_window:
                main_window.update_queue.put(lambda: main_window.log_message(f"\n게임 루프 오류 발생: {error_msg}"))
            is_ready_to_start = False
            time.sleep(1)

def calculate_cell_size(game_region: tuple) -> tuple[float, float]:
    """게임 영역 크기를 기반으로 셀 크기 계산"""
    width, height = game_region[2], game_region[3]
    cell_width = width / NUM_COLS
    cell_height = height / NUM_ROWS
    if main_window:
        main_window.update_gui(lambda: main_window.log_message(f"각 셀의 크기: 가로={cell_width:.1f}px, 세로={cell_height:.1f}px"))
    return (cell_width, cell_height)

def get_button_position(button_name: str) -> tuple[int, int]:
    """사용자가 마우스로 버튼 위치를 지정"""
    log_message(f"\n=== {button_name} 버튼 위치 지정 ===")
    log_message(f"{button_name} 버튼 위에 마우스를 올려주세요")
    log_message("5초 후에 위치를 확인합니다...")
    time.sleep(5)
    x, y = position()
    log_message(f"{button_name} 버튼 위치: ({x}, {y})")
    return (x, y)

def get_game_region():
    """투명 창으로 게임 영역을 선택"""
    log_message("\n=== 게임 영역 선택 ===")
    log_message("1. 투명한 창을 게임 영역에 맞추세요")
    log_message("2. 창 크기와 위치를 조절하여 게임 영역을 정확히 덮어주세요")
    log_message("3. 완료되면 'Space' 키를 누르거나 '영역 선택 완료' 버튼을 클릭하세요")
    
    window = TransparentWindow()
    region = window.get_region()
    
    if region:
        log_message(f"\n선택된 게임 영역:")
        log_message(f"위치: ({region[0]}, {region[1]})")
        log_message(f"크기: {region[2]}x{region[3]}")
        return region
    return None

def wait_for_start():
    """시작 신호를 기다림"""
    global running, is_ready_to_start
    log_message("\n게임을 시작할 준비가 되면 'F2' 키를 눌러주세요...")
    # 키 체크 스레드 시작
    key_thread = threading.Thread(target=check_keys, daemon=True)
    key_thread.start()
    
    while not is_ready_to_start and running:
        time.sleep(0.1)

def resource_path(relative_path):
    """실행 파일 내부의 리소스 경로를 가져옵니다."""
    try:
        # PyInstaller가 생성한 _MEIPASS 경로를 사용
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def find_connected_numbers(grid, row, col, number, visited):
    """DFS로 연결된 같은 숫자들을 찾음"""
    if (row < 0 or row >= NUM_ROWS or 
        col < 0 or col >= NUM_COLS or 
        visited[row][col] or 
        grid[row][col] != number):
        return []
    
    visited[row][col] = True
    connected = [(row, col)]
    
    # 8방향 탐색
    directions = [(-1,-1), (-1,0), (-1,1), (0,-1), (0,1), (1,-1), (1,0), (1,1)]
    for dx, dy in directions:
        new_row, new_col = row + dx, col + dy
        connected.extend(find_connected_numbers(grid, new_row, new_col, number, visited))
    
    return connected

def get_cell_center(row, col):
    """그리드 셀의 중심 좌표를 반환"""
    cell_width = game_region[2] / NUM_COLS
    cell_height = game_region[3] / NUM_ROWS
    
    center_x = game_region[0] + (col + 0.5) * cell_width
    center_y = game_region[1] + (row + 0.5) * cell_height
    
    return (int(center_x), int(center_y))

if __name__ == "__main__":
    main()
