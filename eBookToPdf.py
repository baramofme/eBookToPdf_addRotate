import os
import sys
import time
import mss
import mss.tools
import pyautogui
import natsort
import shutil

from pynput import mouse
from pynput.keyboard import Key, Controller
from PIL import Image

from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QMainWindow, QVBoxLayout, \
    QHBoxLayout, QSlider

import win32gui, win32api

class DrawOnDesktop:
    def __init__(self):
        self.hwnd = win32gui.GetDesktopWindow()
        self.hdc = win32gui.GetDC(self.hwnd)

    def draw_rect(self, x, y, w, h):
        # 색 지정
        color = win32api.RGB(255, 0, 0) # 빨강

        # 가로 직선
        for i in range(x, x + w):
            win32gui.SetPixel(self.hdc, i, y, color)
            win32gui.SetPixel(self.hdc, i, y + h, color)

        # 세로 직선
        for i in range(y, y + h):
            win32gui.SetPixel(self.hdc, x, i, color)
            win32gui.SetPixel(self.hdc, x + w, i, color)

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()

        self.num = 1
        self.rotate_num = 1
        self.posX1 = 0
        self.posY1 = 0
        self.posX2 = 0
        self.posY2 = 0
        self.rotation_button_posX = 0
        self.rotation_button_posY = 0
        self.image_rotate_angle = 0
        self.total_page = 1
        self.speed = 0.1
        self.region = {}
        self.file_list = []

        # 앱 타이틀
        self.setWindowTitle("eBookToPdf")

        # 버튼 생성
        self.button1 = QPushButton("좌표 위치 클릭")
        self.button2 = QPushButton("좌표 위치 클릭")
        self.button3 = QPushButton("PDF로 만들기")
        self.button3.setFixedSize(QSize(430, 60))
        self.button4 = QPushButton("초기화")
        self.btn_start_drag_rotate_left = QPushButton("좌회전 버튼 드래그")
        self.btn_start_drag_rotate_right = QPushButton("우회전 버튼 드래그")
        self.btn_process_screen_rotate = QPushButton("화면 회전 클릭 시작")
        self.btn_rotate_and_save_image_left = QPushButton("이미지 좌회전 저장")
        self.btn_rotate_and_save_image_right = QPushButton("이미지 우회전 저장")

        # 버튼 클릭 이벤트
        self.button1.clicked.connect(self.좌측상단_좌표_클릭)
        self.button2.clicked.connect(self.우측하단_좌표_클릭)
        self.button3.clicked.connect(self.btn_click)
        self.button4.clicked.connect(self.초기화)
        self.btn_start_drag_rotate_left.clicked.connect(self.좌회전_위치_드래그)
        self.btn_start_drag_rotate_right.clicked.connect(self.우회전_위치_드래그)
        self.btn_process_screen_rotate.clicked.connect(self.process_rotate_btn_click)
        self.btn_rotate_and_save_image_left.clicked.connect(self.이미지_저장시_좌회전_방향)
        self.btn_rotate_and_save_image_right.clicked.connect(self.이미지_저장시_우회전_방향)

        # 속도 slider
        self.speed_slider = QSlider(Qt.Orientation.Horizontal)
        self.speed_slider.setMinimum(1)
        self.speed_slider.setMaximum(20)
        self.speed_slider.setValue(1)
        self.speed_slider.valueChanged.connect(self.속도_변경)

        self.speed_label = QLabel(f'캡쳐 속도: {self.speed:.1f}초')
        self.speed_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        font_speed = self.speed_label.font()
        font_speed.setPointSize(10)
        self.speed_label.setFont(font_speed)

        self.title = QLabel('E-Book PDF 생성기', self)
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font_title = self.title.font()
        font_title.setPointSize(20)
        self.title.setFont(font_title)

        self.stat = QLabel('', self)
        self.stat.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font_stat = self.stat.font()
        font_stat.setPointSize(18)
        font_stat.setBold(True)
        self.stat.setFont(font_stat)

        self.sign = QLabel('Made By EastShine', self)
        self.sign.setAlignment(Qt.AlignmentFlag.AlignRight)
        font_sign = self.stat.font()
        font_sign.setPointSize(10)
        font_sign.setItalic(True)
        self.sign.setFont(font_sign)

        self.label1 = QLabel('이미지 좌측상단 좌표   ==>   ', self)
        self.label1_1 = QLabel('(0, 0)', self)
        self.label2 = QLabel('이미지 우측하단 좌표   ==>   ', self)
        self.label2_1 = QLabel('(0, 0)', self)
        self.label5 = QLabel('좌회전 버튼 좌표   ==>   ', self)
        self.label_start_drag_rotate_left_position_value = QLabel('(0, 0)', self)
        self.label6 = QLabel('우회전 버튼 좌표   ==>   ', self)
        self.label_start_drag_rotate_right_position_value = QLabel('(0, 0)', self)
        self.label3 = QLabel('총 페이지 수                       ', self)
        self.label4 = QLabel('PDF 이름                         ', self)

        self.input1 = QLineEdit()
        self.input1.setPlaceholderText("총 페이지 수를 입력하세요.")

        self.input2 = QLineEdit()
        self.input2.setPlaceholderText("생성할 PDF의 이름을 입력하세요.")

        # Box 설정
        box1 = QHBoxLayout()
        box1.addWidget(self.label1)
        box1.addWidget(self.label1_1)
        box1.addWidget(self.button1)

        box2 = QHBoxLayout()
        box2.addWidget(self.label2)
        box2.addWidget(self.label2_1)
        box2.addWidget(self.button2)

        box3 = QHBoxLayout()
        box3.addWidget(self.label3)
        box3.addWidget(self.input1)

        box4 = QHBoxLayout()
        box4.addWidget(self.label4)
        box4.addWidget(self.input2)

        box5 = QHBoxLayout()
        box5.addWidget(self.speed_label)
        box5.addWidget(self.speed_slider)

        box6 = QHBoxLayout()
        box6.addWidget(self.stat)
        box6.addWidget(self.button4)
        box6.addWidget(self.sign)

        box_drag_rotate_right_btn = QHBoxLayout()
        box_drag_rotate_right_btn.addWidget(self.label5)
        box_drag_rotate_right_btn.addWidget(self.label_start_drag_rotate_left_position_value)
        box_drag_rotate_right_btn.addWidget(self.btn_start_drag_rotate_left)

        box_drag_rotate_left_btn = QHBoxLayout()
        box_drag_rotate_left_btn.addWidget(self.label6)
        box_drag_rotate_left_btn.addWidget(self.label_start_drag_rotate_right_position_value)
        box_drag_rotate_left_btn.addWidget(self.btn_start_drag_rotate_right)

        box9 = QHBoxLayout()
        box9.addWidget(self.btn_process_screen_rotate)

        box_select_rotate_direction = QHBoxLayout()
        box_select_rotate_direction.addWidget(self.btn_rotate_and_save_image_left)
        box_select_rotate_direction.addWidget(self.btn_rotate_and_save_image_right)

        # 레이아웃 설정
        layout = QVBoxLayout()
        layout.addWidget(self.title)
        layout.addStretch(1)
        layout.addLayout(box_drag_rotate_right_btn)
        layout.addStretch(1)
        layout.addLayout(box_drag_rotate_left_btn)
        layout.addLayout(box9)
        layout.addLayout(box_select_rotate_direction)
        layout.addStretch(2)
        layout.addLayout(box1)
        layout.addStretch(1)
        layout.addLayout(box2)
        layout.addStretch(1)
        layout.addLayout(box3)
        layout.addStretch(1)
        layout.addLayout(box4)
        layout.addStretch(4)
        layout.addLayout(box5)
        layout.addLayout(box6)
        layout.addWidget(self.button3)

        container = QWidget()
        container.setLayout(layout)

        self.setCentralWidget(container)

        # 창 크기 고정
        self.setFixedSize(QSize(450, 320))

    def 초기화(self):
        self.num = 1
        self.posX1 = 0
        self.posY1 = 0
        self.posX2 = 0
        self.posY2 = 0
        self.speed = 0.1
        self.total_page = 1
        self.region = {}
        self.label1_1.setText('(0, 0)')
        self.label2_1.setText('(0, 0)')
        self.label_start_drag_rotate_left_position_value.setText('(0, 0)')
        self.label_start_drag_rotate_right_position_value.setText('(0, 0)')
        self.rotation_button_posX = 0
        self.rotation_button_posY = 0
        self.image_rotate_angle = 0
        self.btn_start_drag_rotate_left.setEnabled(True)
        self.btn_start_drag_rotate_right.setEnabled(True)
        self.btn_rotate_and_save_image_left.setEnabled(True)
        self.btn_rotate_and_save_image_right.setEnabled(True)
        self.input1.clear()
        self.input2.clear()
        self.stat.clear()
        self.speed_slider.setValue(1)

    def 좌측상단_좌표_클릭(self):
        def on_click(x, y, button, pressed):
            self.posX1 = int(x)
            self.posY1 = int(y)
            self.label1_1.setText(str(f'({int(x)}, {int(y)})'))
            print('Button: %s, Position: (%s, %s), Pressed: %s ' % (button, x, y, pressed))
            if not pressed:
                return False

        with mouse.Listener(on_click=on_click) as listener:
            listener.join()

    def 우측하단_좌표_클릭(self):
        def on_click(x, y, button, pressed):
            self.posX2 = int(x)
            self.posY2 = int(y)
            self.label2_1.setText(str(f'({int(x)}, {int(y)})'))
            print('Button: %s, Position: (%s, %s), Pressed: %s ' % (button, x, y, pressed))
            if not pressed:
                return False

        with mouse.Listener(on_click=on_click) as listener:
            listener.join()

    def 좌회전_위치_드래그(self):

        def on_move(x, y):
            print('Pointer moved to {0}'.format(
                (x, y)))

        def on_click(x, y, button, pressed):
            nonlocal sX, sY, eX, eY
            print('{0} at {1}'.format(
                'Pressed' if pressed else 'Released',
                (x, y)))
            if pressed:
                sX = x
                sY = y
            elif not pressed:
                eX, eY = x, y
                print(sX, sY, eX, eY)
                DrawOnDesktop().draw_rect(sX, sY, eX-sX, eY-sY)
                self.rotation_button_posX = (int(sX) + int(eX)) / 2
                self.rotation_button_posY = (int(sY) + int(eY)) / 2
                self.label_start_drag_rotate_left_position_value.setText(str(f'({self.rotation_button_posX }, {self.rotation_button_posY})'))

                self.btn_start_drag_rotate_left.setEnabled(True)
                self.btn_start_drag_rotate_right.setEnabled(False)
                self.이미지_저장시_좌회전_방향()
                print(f'Set rotate to left (angle: {self.image_rotate_angle})')
                # Stop listener
                return False

        sX, sY, eX, eY = 0, 0, 0, 0

        # 마우스 리스너 시작
        with mouse.Listener(on_move=on_move,
                            on_click=on_click) as listener:
            listener.join()

    def 우회전_위치_드래그(self):

        def on_move(x, y):
            # print('Pointer moved to {0}'.format(
            #     (x, y)))
            pass

        def on_click(x, y, button, pressed):
            nonlocal sX, sY, eX, eY
            print('{0} at {1}'.format(
                'Pressed' if pressed else 'Released',
                (x, y)))
            if pressed:
                sX = x
                sY = y
            elif not pressed:
                eX, eY = x, y
                print(sX, sY, eX, eY)
                DrawOnDesktop().draw_rect(sX, sY, eX-sX, eY-sY)
                self.rotation_button_posX = (int(sX) + int(eX)) / 2
                self.rotation_button_posY = (int(sY) + int(eY)) / 2
                self.label_start_drag_rotate_right_position_value.setText(str(f'({self.rotation_button_posX }, {self.rotation_button_posY})'))
                self.btn_start_drag_rotate_left.setEnabled(False)
                self.btn_start_drag_rotate_right.setEnabled(True)
                self.이미지_저장시_우회전_방향()
                print(f'Set rotate to left (angle: {self.image_rotate_angle})')
                # Stop listener
                return False

        sX, sY, eX, eY = 0, 0, 0, 0

        # 마우스 리스너 시작
        with mouse.Listener(on_move=on_move,
                            on_click=on_click) as listener:
            listener.join()

    def 이미지_저장시_좌회전_방향(self):

        self.image_rotate_angle = 90

        self.btn_rotate_and_save_image_left.setEnabled(True)
        self.btn_rotate_and_save_image_right.setEnabled(False)

    def 이미지_저장시_우회전_방향(self):

        self.image_rotate_angle = -90

        self.btn_rotate_and_save_image_left.setEnabled(False)
        self.btn_rotate_and_save_image_right.setEnabled(True)



    def 속도_변경(self):
        self.speed = self.speed_slider.value() / 10.0
        self.speed_label.setText(f'캡쳐 속도: {self.speed:.1f}초')

    def process_rotate_btn_click(self):

        try:
            self.total_page = int(self.input1.text())

            pos_x, pos_y = pyautogui.position()
            m = mouse.Controller()
            mouse_left = mouse.Button.left
            kb_control = Controller()

            # 모든 페이지 화면 가로로 미리 회전 하기
            while self.rotate_num <= self.total_page:
                m.position = (self.rotation_button_posX, self.rotation_button_posY)
                m.click(mouse_left)
                time.sleep(0.5)
                m.position = (pos_x, pos_y)
                # time.sleep(0.025)

                # 페이지 넘기기
                kb_control.press(Key.right)
                kb_control.release(Key.right)

                self.rotate_num += 1

            self.rotate_num = 1
            # 첫 페이지로 돌아오기
            while self.rotate_num <= self.total_page:
                # 페이지 넘기기
                kb_control.press(Key.left)
                kb_control.release(Key.left)

                self.rotate_num += 1

        except Exception as e:
            print('예외 발생. ', e)
            self.stat.setText('오류 발생. 종료 후 다시 시도해주세요.')

        finally:
            self.rotate_num = 1

    def btn_click(self):

        if self.input1.text() == '':
            self.stat.setText('페이지 수를 입력하세요.')
            self.input1.setFocus()
            return

        if self.input2.text() == '':
            self.stat.setText('PDF 제목을 입력하세요.')
            self.input2.setFocus()
            return

        pos_x, pos_y = pyautogui.position()

        if not (os.path.isdir('pdf_images')):
            os.mkdir(os.path.join('pdf_images'))

        self.total_page = int(self.input1.text())

        # The screen part to capture
        self.region = {'top': self.posY1, 'left': self.posX1, 'width': self.posX2 - self.posX1,
                       'height': self.posY2 - self.posY1}

        m = mouse.Controller()
        mouse_left = mouse.Button.left
        kb_control = Controller()

        try:
            # 화면 전환 위해 한번 클릭
            time.sleep(2)
            m.position = (self.posX1, self.posY1)

            time.sleep(2)
            m.click(mouse_left)
            time.sleep(2)
            m.position = (pos_x, pos_y)

            # @Todo 첫번째 장이 두 번 캡쳐되어서 저장되는 바람에, 마지막 페이지가 누락됨
            # 저장은 0001, 0002 로 나오는 것으로 보아서, 키보드가 처음에 안눌리나?
            # 파일 저장
            while self.num <= self.total_page:

                time.sleep(self.speed)

                # 캡쳐하기
                with mss.mss() as sct:
                    output_path = f'pdf_images/img_{str(self.num).zfill(4)}.png'
                    # Grab the data
                    sct_img = sct.grab(self.region)
                    if self.image_rotate_angle:
                        # 여기서 이미 RGB 로 변환을 해주는데 나중에 파일 열어서 RGB 로 다시 변환해줄 필요가 있나?
                        img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                        rotated_img = img.rotate(self.image_rotate_angle, expand=True)
                        rotated_img.save(output_path)
                    else:
                        # Save to the picture file
                        mss.tools.to_png(sct_img.rgb, sct_img.size, output=output_path)

                # 페이지 넘기기
                kb_control.press(Key.right)
                kb_control.release(Key.right)

                self.num += 1

            print("캡쳐 완료!")
            self.stat.setText('PDF 변환 중..')
            path = 'pdf_images'
            # 이미지 파일 리스트
            self.file_list = os.listdir(path)
            self.file_list = natsort.natsorted(self.file_list)

            # .DS_Store 파일이름 삭제
            if '.DS_Store' in self.file_list:
                del self.file_list[0]

            img_list = []

            # PDF 첫 페이지 만들어두기
            img_path = 'pdf_images/' + self.file_list[0]
            im_buf = Image.open(img_path)
            cvt_rgb_0 = im_buf.convert('RGB')

            for i in self.file_list:
                img_path = 'pdf_images/' + i
                im_buf = Image.open(img_path)
                cvt_rgb = im_buf.convert('RGB')
                img_list.append(cvt_rgb)

            del img_list[0]

            pdf_name = self.input2.text()
            if pdf_name == '':
                pdf_name = 'default'

            cvt_rgb_0.save(pdf_name + '.pdf', save_all=True, append_images=img_list, quality=100)
            print("PDF 변환 완료!")
            self.stat.setText('PDF 변환 완료!')
            shutil.rmtree('pdf_images/')

        except Exception as e:
            print('예외 발생. ', e)
            self.stat.setText('오류 발생. 종료 후 다시 시도해주세요.')

        finally:
            self.num = 1
            self.file_list = []


app = QApplication(sys.argv)

window = MainWindow()
window.show()

# 이벤트 루프 시작
app.exec()
