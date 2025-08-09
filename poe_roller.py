import tkinter as tk
from tkinter import font, messagebox
import pyautogui
import threading
import time
import keyboard
import sys
import ctypes
import os
import json

# --- 기본 설정 ---
DEFAULT_CHAOS_ORB_POS = (1295, 719)
DEFAULT_CHAOS_CELL_SIZE = 53
GRID_ROWS = 12
GRID_COLS = 12
DEFAULT_GRID_LEFT_BOUNDARY_X = 15
DEFAULT_GRID_RIGHT_BOUNDARY_X = 651
DEFAULT_GRID_TOP_BOUNDARY_Y = 125
DEFAULT_GRID_BOTTOM_BOUNDARY_Y = 761
STOP_KEY_1 = 'esc'
STOP_KEY_2 = 'f10'
CONFIG_FILE = 'poe_roller_config.json'
VERSION = "v1.1"
# -------------

class MapRollerApp:
    def __init__(self, root):
        self.root = root
        self.chaos_pos_center = None
        self.chaos_cell_size = DEFAULT_CHAOS_CELL_SIZE
        self.grid_bounds = {
            'left': DEFAULT_GRID_LEFT_BOUNDARY_X,
            'right': DEFAULT_GRID_RIGHT_BOUNDARY_X,
            'top': DEFAULT_GRID_TOP_BOUNDARY_Y,
            'bottom': DEFAULT_GRID_BOTTOM_BOUNDARY_Y
        }
        self.map_coords = []
        self.is_running = False
        self.automation_thread = None
        self.overlay_windows = []
        self.shift_pressed = False
        self.need_initial_shift = True
        self.setup_mode = False
        self.dragging = None  # 현재 드래그 중인 오버레이
        self.drag_mode = None  # 'move' or 'resize'
        self.drag_start = None
        
        # 설정 로드
        self.load_config()
        
        # 콘솔 창 숨기기
        self.hide_console()

        # UI 설정
        self.setup_ui()
        
        self.generate_coordinates()
        self.setup_hotkeys()

        # pyautogui 설정
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.0
        
        # 종료 시 Shift 키 해제를 위한 이벤트 바인딩
        self.root.protocol("WM_DELETE_WINDOW", self.quit_app)

    def save_config(self):
        """현재 설정을 파일로 저장"""
        config = {
            'chaos_pos': self.chaos_pos_center,
            'chaos_size': self.chaos_cell_size,
            'grid_bounds': self.grid_bounds
        }
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)
        except:
            pass

    def load_config(self):
        """저장된 설정 로드"""
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    self.chaos_pos_center = tuple(config.get('chaos_pos', DEFAULT_CHAOS_ORB_POS))
                    self.chaos_cell_size = config.get('chaos_size', DEFAULT_CHAOS_CELL_SIZE)
                    self.grid_bounds = config.get('grid_bounds', {
                        'left': DEFAULT_GRID_LEFT_BOUNDARY_X,
                        'right': DEFAULT_GRID_RIGHT_BOUNDARY_X,
                        'top': DEFAULT_GRID_TOP_BOUNDARY_Y,
                        'bottom': DEFAULT_GRID_BOTTOM_BOUNDARY_Y
                    })
            else:
                self.chaos_pos_center = DEFAULT_CHAOS_ORB_POS
                self.chaos_cell_size = DEFAULT_CHAOS_CELL_SIZE
        except:
            self.chaos_pos_center = DEFAULT_CHAOS_ORB_POS
            self.chaos_cell_size = DEFAULT_CHAOS_CELL_SIZE

    def hide_console(self):
        """콘솔 창 숨기기"""
        if sys.platform == 'win32':
            console_window = ctypes.windll.kernel32.GetConsoleWindow()
            if console_window:
                ctypes.windll.user32.ShowWindow(console_window, 0)

    def setup_ui(self):
        self.root.title(f"PoE 자동 롤러 - {VERSION}") # <-- 수정된 코드
        self.root.attributes('-alpha', 0.9)
        self.root.attributes('-topmost', True)
        self.root.overrideredirect(True)
        self.root.config(bg='black')

        window_width = 550
        window_height = 280
        screen_width = self.root.winfo_screenwidth()
        x_coordinate = screen_width - window_width - 30
        y_coordinate = 30
        self.root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

        # 상태 표시
        self.status_var = tk.StringVar()
        bold_font = font.Font(family="Malgun Gothic", size=12, weight="bold")
        status_label = tk.Label(self.root, textvariable=self.status_var, fg="lime", bg="black", font=bold_font, wraplength=window_width - 20)
        status_label.pack(pady=5, padx=10)
        
        # 위치 설정 안내
        setup_font = font.Font(family="Malgun Gothic", size=9)
        self.setup_info_var = tk.StringVar()
        setup_info_label = tk.Label(self.root, textvariable=self.setup_info_var, fg="cyan", bg="black", font=setup_font, wraplength=window_width - 20)
        setup_info_label.pack(pady=2)
        
        # 핫키 안내
        hotkey_font = font.Font(family="Malgun Gothic", size=10, weight="bold")
        hotkey_label = tk.Label(self.root, text=f"중지: [{STOP_KEY_1.upper()}] 또는 [{STOP_KEY_2.upper()}]", fg="red", bg="black", font=hotkey_font)
        hotkey_label.pack(pady=2)

        # 속도 선택
        speed_frame = tk.LabelFrame(self.root, text="속도 선택", fg="yellow", bg="black", font=hotkey_font)
        speed_frame.pack(pady=5, padx=10, fill="x")
        
        self.speed_var = tk.StringVar(value="fast")
        tk.Radiobutton(speed_frame, text="느리게", variable=self.speed_var, value="slow", fg="white", bg="black", selectcolor="black", font=hotkey_font).pack(side=tk.LEFT, expand=True)
        tk.Radiobutton(speed_frame, text="보통", variable=self.speed_var, value="normal", fg="white", bg="black", selectcolor="black", font=hotkey_font).pack(side=tk.LEFT, expand=True)
        tk.Radiobutton(speed_frame, text="빠르게", variable=self.speed_var, value="fast", fg="white", bg="black", selectcolor="black", font=hotkey_font).pack(side=tk.LEFT, expand=True)
        tk.Radiobutton(speed_frame, text="극한 속도", variable=self.speed_var, value="max", fg="white", bg="black", selectcolor="black", font=hotkey_font).pack(side=tk.LEFT, expand=True)

        # 버튼 프레임
        button_frame = tk.Frame(self.root, bg="black")
        button_frame.pack(pady=10)

        self.start_button = tk.Button(button_frame, text="시작", command=self.start_automation_thread, width=15, height=2)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.setup_button = tk.Button(button_frame, text="위치 설정", command=self.toggle_setup_mode, width=15, height=2, bg="lightblue")
        self.setup_button.pack(side=tk.LEFT, padx=5)
        
        quit_button = tk.Button(button_frame, text="프로그램 종료", command=self.quit_app, width=15, height=2)
        quit_button.pack(side=tk.LEFT, padx=5)

    def toggle_setup_mode(self):
        """설정 모드 토글"""
        if self.setup_mode:
            self.exit_setup_mode()
        else:
            self.enter_setup_mode()

    def enter_setup_mode(self):
        """설정 모드 진입"""
        self.setup_mode = True
        self.setup_button.config(text="설정 완료", bg="yellow")
        self.start_button.config(state=tk.DISABLED)
        self.setup_info_var.set("드래그로 이동, 우하단 빨간 핸들로 크기 조절")
        
        # 조절 가능한 오버레이 생성
        self.create_adjustable_overlays()

    def exit_setup_mode(self):
        """설정 모드 종료"""
        self.setup_mode = False
        self.setup_button.config(text="위치 설정", bg="lightblue")
        self.start_button.config(state=tk.NORMAL)
        self.setup_info_var.set("")
        
        # 설정 저장
        self.save_config()
        
        # 좌표 재생성
        self.generate_coordinates()
        
        # 오버레이 제거
        self.destroy_visual_overlays()
        
        self.status_var.set(f"설정 완료! 좌표 {len(self.map_coords)}개 생성됨")

    def create_adjustable_overlays(self):
        """조절 가능한 오버레이 생성"""
        self.destroy_visual_overlays()
        
        # 카오스 오브 오버레이
        self.chaos_overlay = self.create_draggable_overlay(
            self.chaos_pos_center[0] - self.chaos_cell_size // 2,
            self.chaos_pos_center[1] - self.chaos_cell_size // 2,
            self.chaos_cell_size,
            self.chaos_cell_size,
            'yellow',
            'chaos'
        )
        
        # 맵 인벤토리 오버레이
        self.grid_overlay = self.create_draggable_overlay(
            self.grid_bounds['left'],
            self.grid_bounds['top'],
            self.grid_bounds['right'] - self.grid_bounds['left'],
            self.grid_bounds['bottom'] - self.grid_bounds['top'],
            'cyan',
            'grid'
        )

    def create_draggable_overlay(self, x, y, width, height, color, overlay_type):
        """드래그 가능한 오버레이 생성"""
        overlay = tk.Toplevel(self.root)
        overlay.attributes('-alpha', 0.3)
        overlay.attributes('-topmost', True)
        overlay.overrideredirect(True)
        overlay.geometry(f"{width}x{height}+{x}+{y}")
        overlay.config(bg=color)
        overlay.overlay_type = overlay_type  # 타입 저장
        
        # 크기 조절용 핸들 (우하단 모서리)
        resize_handle = tk.Frame(overlay, bg='red', width=15, height=15, cursor="sizing")
        resize_handle.place(relx=1.0, rely=1.0, anchor='se')
        
        # 라벨 추가 (어떤 영역인지 표시)
        label_text = "카오스 오브" if overlay_type == 'chaos' else "맵 인벤토리"
        label = tk.Label(overlay, text=label_text, bg=color, fg='black', font=('Arial', 10, 'bold'))
        label.pack(pady=5)
        
        # 이동용 이벤트 바인딩
        overlay.bind('<Button-1>', lambda e: self.start_drag(e, overlay, 'move'))
        overlay.bind('<B1-Motion>', lambda e: self.on_drag(e, overlay))
        overlay.bind('<ButtonRelease-1>', lambda e: self.stop_drag(overlay))
        
        # 크기 조절용 이벤트 바인딩
        resize_handle.bind('<Button-1>', lambda e: self.start_drag(e, overlay, 'resize'))
        resize_handle.bind('<B1-Motion>', lambda e: self.on_drag(e, overlay))
        resize_handle.bind('<ButtonRelease-1>', lambda e: self.stop_drag(overlay))
        
        self.overlay_windows.append(overlay)
        return overlay

    def start_drag(self, event, overlay, mode):
        """드래그 시작"""
        self.dragging = overlay
        self.drag_mode = mode
        
        # 현재 윈도우 정보 저장
        self.drag_start = {
            'mouse_x': event.x_root,
            'mouse_y': event.y_root,
            'window_x': overlay.winfo_x(),
            'window_y': overlay.winfo_y(),
            'window_width': overlay.winfo_width(),
            'window_height': overlay.winfo_height()
        }
        
        # 이벤트 전파 중지
        return "break"

    def on_drag(self, event, overlay):
        """드래그 중"""
        if self.dragging != overlay:
            return
            
        if self.drag_mode == 'move':
            # 윈도우 이동
            dx = event.x_root - self.drag_start['mouse_x']
            dy = event.y_root - self.drag_start['mouse_y']
            new_x = self.drag_start['window_x'] + dx
            new_y = self.drag_start['window_y'] + dy
            
            overlay.geometry(f"{self.drag_start['window_width']}x{self.drag_start['window_height']}+{new_x}+{new_y}")
            
        elif self.drag_mode == 'resize':
            # 크기 조절
            dx = event.x_root - self.drag_start['mouse_x']
            dy = event.y_root - self.drag_start['mouse_y']
            
            new_width = max(50, self.drag_start['window_width'] + dx)
            new_height = max(50, self.drag_start['window_height'] + dy)
            
            overlay.geometry(f"{new_width}x{new_height}+{self.drag_start['window_x']}+{self.drag_start['window_y']}")

    def stop_drag(self, overlay):
        """드래그 종료"""
        if self.dragging != overlay:
            return
            
        # 최종 위치/크기 저장
        if overlay.overlay_type == 'chaos':
            self.chaos_pos_center = (
                overlay.winfo_x() + overlay.winfo_width() // 2,
                overlay.winfo_y() + overlay.winfo_height() // 2
            )
            self.chaos_cell_size = min(overlay.winfo_width(), overlay.winfo_height())
        else:  # grid
            self.grid_bounds = {
                'left': overlay.winfo_x(),
                'right': overlay.winfo_x() + overlay.winfo_width(),
                'top': overlay.winfo_y(),
                'bottom': overlay.winfo_y() + overlay.winfo_height()
            }
        
        self.dragging = None
        self.drag_mode = None

    def generate_coordinates(self):
        """좌표 생성"""
        self.map_coords = []
        
        if not self.chaos_pos_center:
            return
            
        cell_width = (self.grid_bounds['right'] - self.grid_bounds['left']) / GRID_COLS
        cell_height = (self.grid_bounds['bottom'] - self.grid_bounds['top']) / GRID_ROWS
        start_x = self.grid_bounds['left'] + (cell_width / 2)
        start_y = self.grid_bounds['top'] + (cell_height / 2)
        
        for col in range(GRID_COLS):
            for row in range(GRID_ROWS):
                x = round(start_x + (col * cell_width))
                y = round(start_y + (row * cell_height))
                self.map_coords.append((x, y))

    def destroy_visual_overlays(self):
        """오버레이 제거"""
        for window in self.overlay_windows:
            try:
                window.destroy()
            except:
                pass
        self.overlay_windows.clear()

    def setup_hotkeys(self):
        keyboard.add_hotkey(STOP_KEY_1, self.stop_automation)
        keyboard.add_hotkey(STOP_KEY_2, self.stop_automation)

    def start_automation_thread(self):
        if not self.is_running and not self.setup_mode:
            if not self.map_coords:
                messagebox.showwarning("경고", "먼저 위치 설정을 해주세요!")
                return
                
            self.is_running = True
            self.need_initial_shift = True
            self.start_button.config(state=tk.DISABLED)
            self.destroy_visual_overlays()
            
            if self.automation_thread and self.automation_thread.is_alive():
                self.automation_thread.join(timeout=0.5)
                
            self.automation_thread = threading.Thread(target=self.run_automation, daemon=True)
            self.automation_thread.start()

    def stop_automation(self):
        if self.is_running:
            self.is_running = False
            # 즉시 Shift 키 해제
            self.force_release_shift()
            self.root.after(0, self.status_var.set, "중지 신호 감지! 작업을 멈춥니다...")
            self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))

    def force_release_shift(self):
        """강제로 Shift 키 해제"""
        if self.shift_pressed:
            try:
                # 여러 번 시도하여 확실히 해제
                for _ in range(3):
                    pyautogui.keyUp('shift')
                    time.sleep(0.01)
            except:
                pass
            finally:
                self.shift_pressed = False

    def release_shift_key(self):
        """일반적인 Shift 키 해제"""
        if self.shift_pressed:
            try:
                pyautogui.keyUp('shift')
            except:
                pass
            self.shift_pressed = False

    def run_automation(self):
        speed = self.speed_var.get()
        
        # 속도 설정
        if speed == "slow":
            move_duration = 0.15
            click_delay = 0.15
        elif speed == "normal":
            move_duration = 0.08
            click_delay = 0.08
        elif speed == "fast":
            move_duration = 0.03
            click_delay = 0.03
        elif speed == "max":
            move_duration = 0.0
            click_delay = 0.0
            
        try:
            # 초기 카오스 오브 선택
            if self.need_initial_shift:
                self.root.after(0, self.status_var.set, "카오스 오브 선택 중...")
                self.release_shift_key()
                
                pyautogui.keyDown('shift')
                self.shift_pressed = True
                time.sleep(0.05)
                pyautogui.moveTo(self.chaos_pos_center, duration=move_duration)
                pyautogui.rightClick()
                time.sleep(0.1)
                self.need_initial_shift = False

            # 맵 롤링
            self.root.after(0, self.status_var.set, "연속 롤링 시작! (중지: ESC 또는 F10)")
            
            for i, map_pos in enumerate(self.map_coords):
                if not self.is_running:
                    break
                
                if keyboard.is_pressed(STOP_KEY_1) or keyboard.is_pressed(STOP_KEY_2):
                    self.root.after(0, self.status_var.set, "중지 키 감지! 작업을 멈춥니다...")
                    break
                
                self.root.after(0, self.status_var.set, f"롤링 중... [{i+1}/{len(self.map_coords)}]")

                # 안정적인 클릭을 위해 mouseDown, sleep, mouseUp으로 변경**
                click_press_delay = 0.02 # 마우스를 누르고 떼는 사이의 미세한 딜레이**

                if speed == "max":
                    pyautogui.moveTo(map_pos, _pause=False)
                    pyautogui.mouseDown(_pause=False)
                    time.sleep(click_press_delay)
                    pyautogui.mouseUp(_pause=False)
                else:
                    pyautogui.moveTo(map_pos, duration=move_duration)
                    pyautogui.mouseDown()
                    time.sleep(click_press_delay)
                    pyautogui.mouseUp()
                    if click_delay > 0:
                        time.sleep(click_delay)
            
        except pyautogui.FailSafeException:
            self.root.after(0, self.status_var.set, "비상 정지! (마우스가 화면 모서리로 이동됨)")
            # FailSafe 발생 시에도 Shift 키 해제
            self.force_release_shift()
        except Exception as e:
            self.root.after(0, self.status_var.set, f"오류 발생: {str(e)}")
            self.force_release_shift()
        finally:
            # 작업 종료 시 확실하게 Shift 키 해제
            self.force_release_shift()
            self.is_running = False
            self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.status_var.set("준비 완료. 시작 버튼을 눌러 새 작업을 시작하세요."))

    def quit_app(self):
        """프로그램 종료"""
        # 종료 전 Shift 키 확실히 해제
        self.stop_automation()
        self.force_release_shift()
        time.sleep(0.2)
        
        # 한 번 더 확실히 해제
        try:
            for _ in range(3):
                pyautogui.keyUp('shift')
                time.sleep(0.01)
        except:
            pass
            
        keyboard.unhook_all()
        self.destroy_visual_overlays()
        self.root.destroy()
        os._exit(0)

if __name__ == "__main__":
    if ctypes.windll.shell32.IsUserAnAdmin():
        root = tk.Tk()
        app = MapRollerApp(root)
        
        # 초기 안내 메시지
        app.status_var.set("'위치 설정' 버튼을 눌러 영역을 조절하세요")
        
        root.mainloop()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)