import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QDesktopWidget, QLabel, QSpacerItem, QSizePolicy
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QIcon, QPalette, QBrush
from ImageConversionTools.ImportingPicturesIntoWord import App

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        # 设置窗口标题和大小
        self.setWindowTitle("初始化界面")
        self.setFixedSize(800, 600)

        # 设置窗口图标
        self.setWindowIcon(QIcon("./icon/00002.png"))  # 将 "icon.png" 替换为你的图标文件路径

        # 设置背景图片
        self.set_background("./Pictures/00065-3853250584.png")  # 将 "background.jpg" 替换为你的图片路径

        # 创建按钮
        self.btn_convert_to_word = QPushButton("图片文件夹转换为Word")
        self.btn_convert_to_excel = QPushButton("图片文件夹转换为Excel")
        self.btn_compress_image = QPushButton("压缩图片")
        self.btn_crop_image = QPushButton("裁剪图片")
        self.btn_ratio_converter = QPushButton("比例转换器")
        self.btn_exit = QPushButton("退出程序")  # 退出按钮

        # 设置鼠标悬浮提示
        self.btn_compress_image.setToolTip("通过压缩转换图片比例，注意目标图片尺寸不可以大于实际图片尺寸")
        self.btn_crop_image.setToolTip("通过裁剪转换图片比例，注意目标图片尺寸不可以大于实际图片尺寸")

        # 设置按钮样式
        self.btn_convert_to_word.setFixedSize(150, 40)
        self.btn_convert_to_excel.setFixedSize(150, 40)
        self.btn_compress_image.setFixedSize(150, 40)
        self.btn_crop_image.setFixedSize(150, 40)
        self.btn_ratio_converter.setFixedSize(150, 40)
        self.btn_exit.setFixedSize(150, 40)  # 设置退出按钮大小

        # 设置按钮点击事件
        self.btn_convert_to_word.clicked.connect(self.open_ITW_window)
        self.btn_exit.clicked.connect(self.exit_program)  # 绑定退出事件

        # 创建布局
        layout = QVBoxLayout()

        # 第一行：按钮1和按钮2
        row1_layout = QHBoxLayout()
        row1_layout.addWidget(self.btn_convert_to_word)
        row1_layout.addSpacing(30)  # 设置按钮之间的间距
        row1_layout.addWidget(self.btn_convert_to_excel)
        row1_layout.setAlignment(Qt.AlignCenter)

        # 第二行：按钮3和按钮4
        row2_layout = QHBoxLayout()
        row2_layout.addWidget(self.btn_compress_image)
        row2_layout.addSpacing(30)  # 设置按钮之间的间距
        row2_layout.addWidget(self.btn_crop_image)
        row2_layout.setAlignment(Qt.AlignCenter)

        # 第三行：按钮5
        row3_layout = QHBoxLayout()
        row3_layout.addWidget(self.btn_ratio_converter)
        row3_layout.setAlignment(Qt.AlignCenter)

        # 添加布局到主界面
        layout.addLayout(row1_layout)
        layout.addLayout(row2_layout)
        layout.addLayout(row3_layout)

        # 添加版权信息和退出按钮到底部
        footer_layout = QHBoxLayout()
        self.footer_label = QLabel("Copyright:MiemieY|2025.01.12")  # 创建版权信息
        self.footer_label.setAlignment(Qt.AlignLeft | Qt.AlignBottom)
        footer_layout.addWidget(self.footer_label)  # 左下角添加版权信息
        footer_layout.addStretch()  # 中间空白
        footer_layout.addWidget(self.btn_exit)  # 右下角添加退出按钮

        # 添加布局到主界面
        layout.addLayout(footer_layout)

        # 设置主窗口的布局
        self.setLayout(layout)

        # 将窗口居中显示
        self.center_window()

        # 调用样式
        self.Yangshi()

    def Yangshi(self):
        # 设置按钮的样式
        self.btn_convert_to_word.setStyleSheet('''
                   QPushButton {
                       background-color: rgba(255, 192, 203, 0.5);  /* 设置背景颜色为粉色 */
                       color: white;               /* 设置字体颜色为白色 */
                       font-weight: bold;          /* 设置字体加粗 */
                       border-radius: 10px;        /* 设置圆角 */
                   }
                   QPushButton:hover {
                       background-color: #FF1493;  /* 当鼠标悬停时改变背景颜色为玫瑰红 */
                   }
               ''')

        self.btn_convert_to_excel.setStyleSheet('''
                           QPushButton {
                               background-color: rgba(255, 192, 203, 0.5);  /* 设置背景颜色为粉色 */
                               color: white;               /* 设置字体颜色为白色 */
                               font-weight: bold;          /* 设置字体加粗 */
                               border-radius: 10px;        /* 设置圆角 */
                           }
                           QPushButton:hover {
                               background-color: #FF1493;  /* 当鼠标悬停时改变背景颜色为玫瑰红 */
                           }
                       ''')

        self.btn_compress_image.setStyleSheet('''
                           QPushButton {
                               background-color: rgba(255, 192, 203, 0.5);  /* 设置背景颜色为粉色 */
                               color: white;               /* 设置字体颜色为白色 */
                               font-weight: bold;          /* 设置字体加粗 */
                               border-radius: 10px;        /* 设置圆角 */
                           }
                           QPushButton:hover {
                               background-color: #FF1493;  /* 当鼠标悬停时改变背景颜色为玫瑰红 */
                           }
                       ''')

        self.btn_crop_image.setStyleSheet('''
                           QPushButton {
                               background-color: rgba(255, 192, 203, 0.5);  /* 设置背景颜色为粉色 */
                               color: white;               /* 设置字体颜色为白色 */
                               font-weight: bold;          /* 设置字体加粗 */
                               border-radius: 10px;        /* 设置圆角 */
                           }
                           QPushButton:hover {
                               background-color: #FF1493;  /* 当鼠标悬停时改变背景颜色为玫瑰红 */
                           }
                       ''')

        self.btn_ratio_converter.setStyleSheet('''
                           QPushButton {
                               background-color: rgba(255, 192, 203, 0.5);  /* 设置背景颜色为粉色 */
                               color: white;               /* 设置字体颜色为白色 */
                               font-weight: bold;          /* 设置字体加粗 */
                               border-radius: 10px;        /* 设置圆角 */
                           }
                           QPushButton:hover {
                               background-color: #FF1493;  /* 当鼠标悬停时改变背景颜色为玫瑰红 */
                           }
                       ''')

        self.btn_exit.setStyleSheet('''
                                   QPushButton {
                                       background-color: rgba(255, 255, 255, 0.3);  /* 设置背景颜色为粉色 */
                                       color: black;               /* 设置字体颜色为白色 */
                                       font-weight: bold;          /* 设置字体加粗 */
                                       border-radius: 10px;        /* 设置圆角 */
                                   }
                                   QPushButton:hover {
                                       background-color: #FF1493;  /* 当鼠标悬停时改变背景颜色为玫瑰红 */
                                   }
                               ''')

    def center_window(self):
        # 获取屏幕的尺寸
        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()

        # 获取窗口的尺寸
        window_geometry = self.frameGeometry()
        window_width = window_geometry.width()
        window_height = window_geometry.height()

        # 计算窗口的位置，确保它在屏幕的中心
        x_pos = (screen_geometry.width() - window_width) // 2
        y_pos = (screen_geometry.height() - window_height) // 2

        # 设置窗口的位置
        self.move(x_pos, y_pos)

    def set_background(self, image_path):
        # 设置背景图片
        palette = QPalette()
        pixmap = QPixmap(image_path)  # 加载背景图片
        pixmap = pixmap.scaled(self.size(), Qt.KeepAspectRatioByExpanding)  # 缩放图片以填充窗口
        palette.setBrush(QPalette.Background, QBrush(pixmap))
        self.setPalette(palette)

    def exit_program(self):
        # 退出程序
        QApplication.quit()

    def open_ITW_window(self):
        # 创建并启动 ITW 的应用窗口
        self.itw_app = App()  # 实例化 ITW 中的 App 类
        self.itw_app.show()  # 显示 ITW 的窗口
        # 注意：这里不关闭主窗口，保留原UI界面

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
