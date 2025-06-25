import sys  # 系统相关操作，如退出应用
import re  # 正则表达式处理
import numpy as np  # 数值计算库
from skimage.measure import marching_cubes
# import sympy  # 符号计算库
# from memory_profiler import profile
from sympy.printing import pretty  # 美化符号表达式输出
from PyQt5.QtGui import QInputMethod, QFont  # 输入法支持
from PyQt5.QtWidgets import (  # PyQt5 GUI组件
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QPushButton, QScrollArea, QLineEdit, QCheckBox, QSplitter,
    QLabel, QDockWidget, QSizePolicy, QToolBar, QToolButton, QStackedWidget, QTabWidget, QComboBox, QSpinBox, QSlider,
    QColorDialog, QMenu, QMessageBox
)
from PyQt5.QtCore import Qt, QTimer  # Qt核心模块和定时器
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas  # Matplotlib与Qt集成
from matplotlib.figure import Figure  # Matplotlib绘图基础
from sympy import lambdify, symbols, sympify, singularities, S, degree, solve, FiniteSet, solveset, latex, diff, \
    EmptySet, limit, oo, Eq, idiff  # SymPy符号计算相关函数
from sympy.calculus.util import continuous_domain  # 计算函数连续性
from sympy import Pow, Rational
from mpl_toolkits.mplot3d import Axes3D



def type_num(list_exp):
    """
    判断输入表达式列表的类型（显式函数或隐式方程）
    :param list_exp: 表达式字符列表（如包含'x'、'y'或'='）
    :return: 类型代码（1为显式函数，3为隐式方程）
    """
    if list_exp is None:
        return 0
    else:
        # 检查是否包含特定字符
        num_list = [1 if i in list_exp else 0 for i in ["x", "y", "=", "z"]]
        total = sum(num_list)
        if "z=" in "".join(list_exp) and ("x" in list_exp or "y" in list_exp):
            return 4  # 三维曲面
        elif total > 3:
            return 5
        elif total == 1 and num_list[0] == 1:
            return 1
        elif total == 3:
            return 3
        else:
            return 0


# 修改setup_3d_axes函数中的标签和刻度设置：
def setup_3d_axes(ax):
    """统一设置三维坐标轴样式（包含刻度和标签）"""
    axis_range = 3
    tick_interval = 0.5  # 刻度间隔
    tick_length = 0.1  # 刻度线长度

    ax.set_xlim(-axis_range, axis_range)
    ax.set_ylim(-axis_range, axis_range)
    ax.set_zlim(-axis_range, axis_range)

    # 关闭默认坐标轴和背景
    ax.set_axis_off()
    ax.grid(False)
    ax.xaxis.pane.fill = ax.yaxis.pane.fill = ax.zaxis.pane.fill = False


    # 启用剪裁功能
    ax.set_clip_box = True

    # 绘制坐标轴线
    ax.plot([-axis_range, axis_range], [0, 0], [0, 0], color='k', lw=1)  # X轴
    ax.plot([0, 0], [-axis_range, axis_range], [0, 0], color='k', lw=1)  # Y轴
    ax.plot([0, 0], [0, 0], [-axis_range, axis_range], color='k', lw=1)  # Z轴

    # 生成刻度位置（-1.5, -1.0, ..., 1.5）
    tick_positions = np.arange(-axis_range, axis_range + tick_interval, tick_interval)

    # 绘制X轴刻度
    for x in tick_positions:
        if abs(x) < 0.1: continue  # 跳过原点附近
        # Y方向刻度线
        ax.plot([x, x], [-tick_length, tick_length], [0, 0], color='k', lw=0.8)
        # Z方向刻度线
        # ax.plot([x, x], [0, 0], [-tick_length, tick_length], color='k', lw=0.8)
        # 标签（偏移避免重叠）
        ax.text(x, -tick_length * 2, -tick_length * 2, f"{-x:.1f}",
                ha='center', va='top', fontsize=8, color='k')

    # 绘制Y轴刻度
    for y in tick_positions:
        if abs(y) < 0.1: continue
        # X方向刻度线
        ax.plot([-tick_length, tick_length], [y, y], [0, 0], color='k', lw=0.8)
        # Z方向刻度线
        # ax.plot([0, 0], [y, y], [-tick_length, tick_length], color='k', lw=0.8)
        ax.text(-tick_length * 2, y, -tick_length * 2, f"{-y:.1f}",
                ha='right', va='center', fontsize=8, color='k')

    # 绘制Z轴刻度
    for z in tick_positions:
        if abs(z) < 0.1: continue
        # X方向刻度线
        ax.plot([-tick_length, tick_length], [0, 0], [z, z], color='k', lw=0.8)
        # Y方向刻度线
        # ax.plot([0, 0], [-tick_length, tick_length], [z, z], color='k', lw=0.8)
        ax.text(-tick_length * 2, -tick_length * 2, z, f"{z:.1f}",
                ha='right', va='center', fontsize=8, color='k')

    # 轴标签（调整位置）
    label_offset = axis_range * 1.1
    ax.text(-label_offset, 0, 0, 'X', ha='left', va='center', fontsize=10)
    ax.text(0, -label_offset, 0, 'Y', ha='center', va='bottom', fontsize=10)
    ax.text(0, 0, label_offset, 'Z', ha='center', va='bottom', fontsize=10)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("函数图像绘制")  # 窗口标题
        self.setGeometry(100, 100, 1100, 900)  # 窗口位置和大小

        # 存储绘图对象、定时器、当前属性和选中线条
        self.plotted_lines = {}  # 键：输入框对象，值：Matplotlib线条对象
        self.timers = {}  # 键：输入框对象，值：QTimer对象
        self.current_properties = {}  # 存储当前线条属性（颜色、线型等）
        self.selected_line = None  # 当前选中的线条

        # 主布局分割器（水平方向）
        splitter = QSplitter(Qt.Horizontal)

        # 左侧复合区域
        self.left_widget = self.create_left_widget()
        splitter.addWidget(self.left_widget)

        # 修改中间区域创建逻辑
        self.center_widget = self.create_center_widget()
        splitter.addWidget(self.center_widget)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        splitter.setSizes([300, 900])

        # 浮动按钮
        self.expand_button = QPushButton("展开")
        self.expand_button.setFixedSize(80, 40)
        self.expand_button.clicked.connect(self.toggle_left_panel)
        self.expand_button.setVisible(False)
        self.expand_button.setWindowFlag(Qt.FramelessWindowHint)  # 去除窗口边框
        self.expand_button.setWindowFlag(Qt.Tool)  # 设置为工具窗口，不影响主窗口布局
        self.expand_button.setParent(self.center_widget)
        self.expand_button.move(20, 20)
        self.expand_button.setStyleSheet("background-color: lightgray; border-radius: 5px;")

        self.add_item()
        self.splitter = splitter
        self.setCentralWidget(splitter)

    def create_left_widget(self):
        """创建左侧复合布局（可调整的上下结构）"""
        # 创建垂直分割器
        splitter1 = QSplitter(Qt.Vertical)
        splitter1.setHandleWidth(12)  # 设置分割条宽度

        splitter1.setChildrenCollapsible(True)

        # 上半部分：输入框区域
        input_widget = QWidget()
        input_layout = QVBoxLayout(input_widget)
        input_layout.setContentsMargins(10, 10, 10, 10)

        # 折叠按钮和添加按钮
        self.collapse_button = QPushButton("隐藏")
        self.collapse_button.clicked.connect(self.toggle_left_panel)
        self.collapse_button.setFont(QFont("SimHei", 14))
        self.add_button = QPushButton("添加")
        self.add_button.clicked.connect(self.add_item)
        self.add_button.setFont(QFont("SimHei", 14))

        input_layout.addWidget(self.collapse_button)
        input_layout.addWidget(self.add_button)

        # 滚动区域
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_area_content)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll_layout.setSpacing(1)
        self.scroll_layout.setAlignment(Qt.AlignTop)  # 左上角对齐
        self.scroll_area.setWidget(self.scroll_area_content)
        input_layout.addWidget(self.scroll_area)

        # 下半部分：属性面板
        pro_widget = QWidget()
        pro_layout = QVBoxLayout(pro_widget)
        input_layout.setContentsMargins(10, 10, 10, 10)
        self.properties_tabs = self.create_properties_tabs()
        pro_layout.addWidget(self.properties_tabs)

        # 将两个部分添加到分割器
        splitter1.addWidget(input_widget)
        splitter1.addWidget(pro_widget)

        # 设置初始比例和拉伸因子
        splitter1.setSizes([350, 350])  # 初始高度比例
        splitter1.setStretchFactor(0, 2)  # 输入区域拉伸系数
        splitter1.setStretchFactor(1, 2)  # 属性区域拉伸系数

        self.collapsed = False  # 用于判断当前左侧面板是否折叠

        return splitter1

    def create_properties_tabs(self):
        """创建属性选项卡"""
        tab_widget = QTabWidget()
        tab_widget.setTabPosition(QTabWidget.North)

        # 图形属性页
        properties_widget = QWidget()
        properties_layout = QVBoxLayout(properties_widget)
        properties_layout.setSpacing(10)

        # 颜色选择
        color_row = QHBoxLayout()
        color_label = QLabel("颜色:")
        color_label.setFont(QFont("SimSun", 14))
        self.color_button = QPushButton()
        self.color_button.setStyleSheet("background-color: black;")
        color_row.addWidget(color_label)
        color_row.addWidget(self.color_button)
        self.color_button.clicked.connect(self.select_color)
        properties_layout.addLayout(color_row)

        # 线型选择
        line_style_group = QVBoxLayout()
        line_style_label = QLabel("线型:")
        line_style_label.setFont(QFont("SimSun", 14))
        self.line_style_combo = QComboBox()
        self.line_style_combo.addItems(["-", "--", "-.", ":"])
        self.line_style_combo.currentIndexChanged.connect(self.apply_changes_to_line)
        line_style_group.addWidget(line_style_label)
        line_style_group.addWidget(self.line_style_combo)
        properties_layout.addLayout(line_style_group)

        # 线宽设置
        line_width_group = QVBoxLayout()
        line_width_label = QLabel("线宽:")
        line_width_label.setFont(QFont("SimSun", 14))
        self.line_width_spinbox = QSpinBox()
        self.line_width_spinbox.setRange(1, 10)
        self.line_width_spinbox.valueChanged.connect(self.apply_changes_to_line)
        line_width_group.addWidget(line_width_label)
        line_width_group.addWidget(self.line_width_spinbox)
        properties_layout.addLayout(line_width_group)

        # 透明度设置
        alpha_group = QVBoxLayout()
        alpha_label = QLabel("透明度:")
        alpha_label.setFont(QFont("SimSun", 14))
        self.alpha_slider = QSlider(Qt.Horizontal)
        self.alpha_slider.setRange(0, 100)
        self.alpha_slider.setValue(100)
        self.alpha_slider.valueChanged.connect(self.apply_changes_to_line)
        alpha_group.addWidget(alpha_label)
        alpha_group.addWidget(self.alpha_slider)
        properties_layout.addLayout(alpha_group)

        properties_layout.addStretch()

        # 数学属性页
        math_properties_widget = QWidget()
        math_properties_layout = QVBoxLayout(math_properties_widget)
        math_properties_layout.setSpacing(10)

        # 数学属性标签（保持原功能）
        self.domain_label = QLabel("定义域: 未计算")
        self.range_label = QLabel("值域: 未计算")
        self.parity_label = QLabel("对称性: 未计算")
        self.monotonicity_label = QLabel("导数特性: 未计算")
        self.extrema_label = QLabel("极值点: 未计算")
        self.zeros_label = QLabel("解集: 未计算")
        self.asymptotes_label = QLabel("渐近线: 未计算")

        math_properties_layout.addWidget(self.domain_label)
        math_properties_layout.addWidget(self.range_label)
        math_properties_layout.addWidget(self.parity_label)
        math_properties_layout.addWidget(self.monotonicity_label)
        math_properties_layout.addWidget(self.extrema_label)
        math_properties_layout.addWidget(self.zeros_label)
        math_properties_layout.addWidget(self.asymptotes_label)
        math_properties_layout.addStretch()

        # 设置选项卡标题字体样式
        tab_widget.setStyleSheet("""
                QTabBar::tab {
                    font: bold 18px 'Microsoft YaHei';  /* 增大字号并使用更现代字体 */
                    min-width: 80px;  /* 最小宽度保证标签不拥挤 */
                    min-height: 20px;  /* 增加标签高度 */
                    padding: 12px 24px;  /* 增大标签内边距 */
                    margin: 4px 2px;
                    border-radius: 4px;
                }

                QTabBar::tab:selected {
                    background: #F0F6FF;  /* 更明显的选中背景 */
                    color: #1A73E8;      /* 强调色文字 */
                }
            """)

        tab_widget.addTab(properties_widget, "图形属性")
        tab_widget.addTab(math_properties_widget, "数学属性")
        return tab_widget

    # @profile
    def add_item(self):
        """动态添加输入框、复选框和删除按钮到滚动区域"""
        item_layout = QHBoxLayout()  # 水平布局
        item_layout.setContentsMargins(5, 5, 5, 5)
        item_layout.setSpacing(1)

        # 复选框（控制显示/隐藏）
        checkbox = QCheckBox("显示")
        checkbox.setFont(QFont("SimSun-bold", 12))
        checkbox.setChecked(True)  # 默认选中
        line_edit = QLineEdit()  # 输入数学表达式
        line_edit.setFont(QFont("SimSun-bold", 13))

        # 设置输入框右键菜单（用于创建参数滑块）
        line_edit.setContextMenuPolicy(Qt.CustomContextMenu)
        line_edit.customContextMenuRequested.connect(lambda pos: self.show_context_menu(line_edit, pos))

        delete_button = QPushButton("删除")  # 删除当前条目
        delete_button.setFont(QFont("SimSun-bold", 13))

        # 设置控件尺寸策略
        checkbox.setFixedWidth(60)
        checkbox.setFixedHeight(40)
        delete_button.setFixedWidth(60)
        delete_button.setFixedHeight(40)
        line_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # 输入框横向扩展
        line_edit.setFixedHeight(40)

        # 创建条目容器
        item_widget = QWidget()
        item_widget.setLayout(item_layout)

        # 连接信号与槽
        delete_button.clicked.connect(lambda: self.delete_item(item_widget, line_edit))
        checkbox.toggled.connect(lambda checked: self.toggle_plot(line_edit, checked))

        # 输入框内容变化时启动定时器（延迟更新绘图）
        timer = QTimer(self)
        timer.setInterval(1000)  # 1秒后触发
        timer.setSingleShot(True)
        self.timers[line_edit] = timer
        line_edit.textChanged.connect(lambda: self.start_timer(timer, line_edit))

        # 将控件添加到布局
        item_layout.addWidget(checkbox)
        item_layout.addWidget(line_edit)
        item_layout.addWidget(delete_button)
        self.scroll_layout.addWidget(item_widget)

        # 默认绘制初始图像（若复选框选中）
        if checkbox.isChecked():
            self.toggle_plot(line_edit, True)

    def show_context_menu(self, line_edit, pos):
        """显示右键菜单（用于创建参数滑块）"""
        context_menu = QMenu(self)
        create_slider_action = context_menu.addAction("创建参数滑块")
        create_slider_action.triggered.connect(lambda: self.create_parameter_sliders(line_edit))
        context_menu.exec_(line_edit.mapToGlobal(pos))  # 在鼠标位置显示菜单

    def create_parameter_sliders(self, line_edit):
        """为输入框中的参数创建滑块控件"""
        expression = line_edit.text()
        # 使用正则表达式提取参数（排除x和y）
        parameters = re.findall(r'\b[a-zA-Z_]\w*\b', expression)
        parameters = [p for p in parameters if p not in ['x', 'y']]

        # 清除旧滑块（避免重复）
        if hasattr(line_edit, 'sliders'):
            for slider_info in line_edit.sliders[:]:
                self.delete_slider(line_edit, slider_info['param'])

        line_edit.sliders = []  # 存储滑块相关信息
        line_edit.parameters = {}  # 存储参数当前值

        for param in parameters:
            slider_layout = QHBoxLayout()
            label = QLabel(f"{param}:")  # 参数标签
            label.setFont(QFont("SimSun-bold", 15))
            slider = QSlider(Qt.Horizontal)  # 水平滑块
            slider.sliderPressed.connect(lambda p=param: setattr(line_edit, 'is_slider_adjusting', True))
            slider.sliderReleased.connect(lambda p=param: setattr(line_edit, 'is_slider_adjusting', False))
            slider.setRange(0, 100)  # 范围0-100（对应0.0-10.0）
            slider.setValue(50)  # 初始值50（对应5.0）
            value_label = QLabel("5.00")  # 显示当前值
            value_label.setFont(QFont("SimSun-bold", 15))

            # 连接滑块值变化事件
            slider.valueChanged.connect(
                lambda value, vl=value_label: self.update_parameter_value_label(vl, value))
            slider.valueChanged.connect(lambda value, p=param: self.update_parameter(line_edit, p, value))

            delete_button = QPushButton("删除滑块")  # 删除当前滑块
            delete_button.clicked.connect(lambda _, p=param: self.delete_slider(line_edit, p))

            # 将控件添加到布局
            slider_layout.addWidget(label)
            slider_layout.addWidget(slider)
            slider_layout.addWidget(value_label)
            slider_layout.addWidget(delete_button)

            # 将滑块布局插入到输入框条目之后
            input_widget = line_edit.parentWidget()
            input_index = self.scroll_layout.indexOf(input_widget)
            self.scroll_layout.insertLayout(input_index + 1, slider_layout)

            # 初始化参数值并存储信息
            line_edit.parameters[param] = slider.value() / 100.0 * 10
            slider_info = {
                'layout': slider_layout,
                'param': param,
                'slider': slider,
                'value_label': value_label,
                'delete_button': delete_button
            }
            line_edit.sliders.append(slider_info)

        self.update_plot(line_edit)  # 立即更新绘图

    def update_parameter_value_label(self, value_label, value):
        """更新参数值标签（格式化为两位小数）"""
        value_label.setText(f"{value / 100.0 * 10:.2f}")

    def delete_slider(self, line_edit, param):
        """删除指定参数的滑块及其所有控件"""
        if not hasattr(line_edit, 'sliders'):
            return

        # 遍历找到对应参数的滑块信息
        for slider_info in line_edit.sliders[:]:
            if slider_info['param'] == param:
                # 递归删除布局内的所有控件
                def delete_layout_items(layout):
                    while layout.count():
                        item = layout.takeAt(0)
                        widget = item.widget()
                        if widget:
                            widget.deleteLater()
                        sub_layout = item.layout()
                        if sub_layout:
                            delete_layout_items(sub_layout)

                # 从滚动布局中移除并删除
                self.scroll_layout.removeItem(slider_info['layout'])
                delete_layout_items(slider_info['layout'])
                slider_info['layout'].deleteLater()
                line_edit.sliders.remove(slider_info)
                break

        # 清理参数数据
        if param in line_edit.parameters:
            del line_edit.parameters[param]
        if not line_edit.sliders:  # 若无滑块则删除属性
            delattr(line_edit, 'sliders')

        self.update_plot(line_edit)  # 更新绘图

    def update_parameter(self, line_edit, param, value):
        """更新参数值并重新绘图"""
        line_edit.parameters[param] = value / 100.0 * 10  # 将滑块值映射到0.0-10.0
        self.update_plot(line_edit)

    def delete_item(self, item_widget, line_edit):
        """删除整个输入框条目（包括关联的滑块和绘图）"""
        # 删除所有关联滑块
        if hasattr(line_edit, 'sliders'):
            for slider_info in line_edit.sliders[:]:
                self.delete_slider(line_edit, slider_info['param'])

        # 从布局中移除并销毁条目
        if self.scroll_layout.count() > 1:
            self.scroll_layout.removeWidget(item_widget)
            item_widget.deleteLater()
        else:
            line_edit.clear()  # 若只剩一个条目则清空输入框

        # 移除对应的绘图
        if line_edit in self.plotted_lines:
            line = self.plotted_lines.pop(line_edit)
            line.remove()
            self.canvas.draw()

    def toggle_plot(self, line_edit, checked):
        """根据复选框状态显示或隐藏绘图"""
        if checked:
            self.update_plot(line_edit)  # 显示
        else:
            if line_edit in self.plotted_lines:
                line = self.plotted_lines.pop(line_edit)
                line.remove()
                self.canvas.draw()

    def create_center_widget(self):
        """创建中间绘图区域（Matplotlib画布）"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)  # 增加画布四周边距
        self.figure = Figure()  # 创建Figure对象
        self.canvas = FigureCanvas(self.figure)  # 将Figure嵌入Qt
        layout.addWidget(self.canvas)
        ax = self.figure.gca()
        # 调整刻度样式
        ax.tick_params(axis='both',
                       which='major',
                       labelsize=18,  # 刻度标签字号
                       labelcolor='#333')  # 标签颜色
        self.canvas.mpl_connect("button_press_event", self.on_canvas_click)  # 绑定点击事件
        self.figure.gca()  # 初始化坐标轴
        self.canvas.draw()
        return widget

    def update_plot(self, line_edit):
        """根据输入框内容更新绘图（显式函数或隐式方程）"""
        try:
            x, y, z = symbols('x y z')
            expression = line_edit.text()

            if not expression:  # 空输入则清除绘图
                if line_edit in self.plotted_lines:
                    line = self.plotted_lines.pop(line_edit)
                    line.remove()
                    self.canvas.draw()
                return

            # 替换表达式中的参数为实际值
            if hasattr(line_edit, 'parameters'):
                for param, value in line_edit.parameters.items():
                    expression = expression.replace(param, str(value))

            kind = type_num(list(expression))  # 判断函数类型
            # self.figure.clf()
            ax = self.figure.gca()  # 获取当前坐标轴
            if kind == 4:  # 三维曲面处理
                # if not isinstance(ax, Axes3D):
                #     self.figure.delaxes(ax)
                #     ax = self.figure.add_subplot(111, projection='3d')
                self.figure.clf()
                ax = self.figure.add_subplot(111, projection='3d')
                ax.set_box_aspect([1, 1, 1])  # 强制等比例缩放
                ax.clear()  # 清除旧图

                axis_range = 3
                plot_range = 12  # 坐标轴绘制范围比数据范围大
                ax.set_xlim(-plot_range, plot_range)
                ax.set_ylim(-plot_range, plot_range)
                ax.set_zlim(-plot_range, plot_range)

                # 调整子图边距，避免绘图区被压缩
                self.figure.subplots_adjust(left=0, right=1, bottom=0, top=1, wspace=0, hspace=0)

                setup_3d_axes(ax)

                _, rhs = expression.split("=")
                expr = sympify(rhs)

                # 生成网格数据（严格限制在坐标轴范围内）
                x_vals = np.linspace(-axis_range, axis_range, 200)
                y_vals = np.linspace(-axis_range, axis_range, 200)
                X, Y = np.meshgrid(x_vals, y_vals)

                # 转换为数值函数
                func = lambdify((x, y), expr, 'numpy')
                Z = func(X, Y)

                # 创建三维掩码（X/Y/Z均不超过范围）
                mask = (np.abs(X) <= axis_range) & \
                       (np.abs(Y) <= axis_range) & \
                       (np.abs(Z) <= axis_range)
                Z = np.where(mask, Z, np.nan)  # 超出部分设为NaN

                # 绘制曲面
                surf = ax.plot_surface(
                    X, Y, Z,
                    cmap='viridis',
                    edgecolor='none',
                    alpha=0.8,
                    rstride=2,
                    cstride=2
                )

                self.plotted_lines[line_edit] = surf

            elif kind == 5:  # 隐式三维方程（如 x² + y² + z² = 1）
                # if not isinstance(ax, Axes3D):
                #     self.figure.delaxes(ax)
                #     ax = self.figure.add_subplot(111, projection='3d')
                self.figure.clf()
                ax = self.figure.add_subplot(111, projection='3d')
                ax.set_box_aspect([1, 1, 1])  # 强制等比例缩放
                ax.clear()  # 清除旧图

                # 调整子图边距，避免绘图区被压缩
                self.figure.subplots_adjust(left=0, right=1, bottom=0, top=1, wspace=0, hspace=0)


                setup_3d_axes(ax)

                # --- 调整网格范围和顶点映射 ---
                # 扩大采样范围以确保完整显示球体
                grid_range = 3  # 略大于方程半径1
                x_vals = np.linspace(-grid_range, grid_range, 40)
                y_vals = np.linspace(-grid_range, grid_range, 40)
                z_vals = np.linspace(-grid_range, grid_range, 40)
                X, Y, Z = np.meshgrid(x_vals, y_vals, z_vals)

                # 解析方程
                lhs, rhs = expression.split("=")
                implicit_func = sympify(lhs) - sympify(rhs)

                # 计算方程值
                f = lambdify((x, y, z), implicit_func, 'numpy')
                F = f(X, Y, Z)

                # 提取等值面（level=0对应方程解）
                verts, faces, _, _ = marching_cubes(F, level=0, spacing=(
                    x_vals[1] - x_vals[0],
                    y_vals[1] - y_vals[0],
                    z_vals[1] - z_vals[0]
                ))

                # 正确映射顶点坐标到实际范围（无需缩放）
                verts = verts + np.array([x_vals[0], y_vals[0], z_vals[0]])

                # 绘制网格
                ax.plot_trisurf(verts[:, 0], verts[:, 1], faces, verts[:, 2], cmap='viridis', alpha=0.6)

                # 存储绘图对象
                self.plotted_lines[line_edit] = ax

            elif kind == 1:  # 显式函数（y = f(x)）
                if isinstance(ax, Axes3D):
                    ax.set_xlim(-3, 3)
                    ax.set_ylim(-3, 3)

                x1 = np.linspace(-5, 5, 1000)  # 生成x值
                func = lambdify(x, expression, 'numpy')  # 转换为可调用函数
                y1 = func(x1)

                # 创建或更新线条
                if line_edit not in self.plotted_lines:
                    line, = ax.plot(x1, y1, color="blue", linestyle="-", linewidth=2, alpha=1.0)
                    self.plotted_lines[line_edit] = line
                else:
                    line = self.plotted_lines[line_edit]
                    line.set_ydata(y1)
                    line.set_label(expression)

                # 自动调整坐标轴（非滑块调整时）
                if not getattr(line_edit, 'is_slider_adjusting', False):
                    ax.relim()
                    ax.autoscale_view()

            elif kind == 3:  # 隐式方程（f(x,y)=0）

                lhs, rhs = expression.split("=")
                implicit_func = sympify(lhs) - sympify(rhs)
                # 转换分数幂：将 x^(a/b) 转换为 (x^a)^(1/b)

                implicit_func = implicit_func.replace(
                    lambda e: e.is_Pow and e.exp.is_Rational,
                    lambda e: Pow(e.base ** e.exp.p, Rational(1, e.exp.q))
                )

                f = lambdify((x, y), implicit_func, 'numpy')  # 二元函数

                # 生成网格点
                if "grid" not in self.plotted_lines:
                    x_vals = np.linspace(-3, 3, 1000)
                    y_vals = np.linspace(-3, 3, 1000)
                    self.plotted_lines["grid"] = np.meshgrid(x_vals, y_vals)

                x_grid, y_grid = self.plotted_lines["grid"]
                z = f(x_grid, y_grid)  # 计算z值

                # 绘制等高线（z=0）
                if line_edit not in self.plotted_lines:
                    contour = ax.contour(x_grid, y_grid, z, levels=[0])
                    self.plotted_lines[line_edit] = contour
                else:
                    # 清除旧等高线并绘制新线
                    for collection in self.plotted_lines[line_edit].collections:
                        collection.remove()
                    contour = ax.contour(x_grid, y_grid, z, levels=[0])
                    self.plotted_lines[line_edit] = contour

            self.canvas.draw()  # 重绘画布

        except Exception as e:
            name = str(e).split("(")
            if name[0][:4] == "name" or name[0][:4] == "Cann":
                pass
            else:
                print(f"绘图错误: {name[0]}")
                QMessageBox.warning(self, "输入错误", f"表达式输入错误：{name[0]}")
                if line_edit in self.plotted_lines:
                    line = self.plotted_lines.pop(line_edit)
                    line.remove()
                    self.canvas.draw()

    def start_timer(self, timer, line_edit):
        """启动定时器以延迟更新绘图（避免频繁重绘）"""
        timer.stop()
        try:
            timer.timeout.disconnect()  # 断开所有已连接的槽函数
        except TypeError:
            pass  # 如果没有连接任何槽函数，忽略错误
        timer.timeout.connect(lambda: self.update_plot(line_edit))
        timer.start()

    def on_canvas_click(self, event):
        """处理画布点击事件（选中线条并更新属性面板）"""
        for line_edit, line in self.plotted_lines.items():
            if isinstance(line, list):  # 跳过网格数据
                continue
            contains, _ = line.contains(event)
            if contains:
                self.selected_line = line  # 记录选中线条
                self.update_properties_panel(line)  # 更新属性面板
                self.update_math_properties_panel(line_edit)  # 更新数学属性
                break

    def update_math_properties_panel(self, line_edit):
        """更新数学属性面板（使用SymPy计算）"""
        expression = line_edit.text()
        x, y = symbols('x y')

        try:
            # 判断是否为显式函数（y = f(x)）
            is_explicit = '=' not in expression
            if is_explicit:
                expr = sympify(expression)
                is_linear = expr.is_polynomial() and degree(expr, x) == 1  # 是否线性
            else:
                lhs, rhs = expression.split('=')
                equation = Eq(sympify(lhs), sympify(rhs))  # 创建方程对象

            # ------------------------- 计算定义域 -------------------------
            if is_explicit:
                if is_linear or (expr.is_polynomial() and degree(expr, x) > 0):
                    domain_str = "(-∞, ∞)"  # 多项式函数定义域为全体实数
                else:
                    domain = continuous_domain(expr, x, S.Reals)  # 连续区间
                    singular = singularities(expr, x, S.Reals)  # 奇点
                    domain_str = self.domain_to_str(domain - singular)  # 格式化为字符串
            else:
                domain_str = "隐函数定义域需具体分析"  # 隐式方程特殊处理

            self.domain_label.setText(f"定义域: {domain_str}")

            # ------------------------- 计算奇偶性/对称性 -------------------------
            if is_explicit:
                f_minus_x = expr.subs(x, -x)  # 替换x为-x
                if expr == f_minus_x:
                    parity = "偶函数"
                elif expr == -f_minus_x:
                    parity = "奇函数"
                else:
                    parity = "非奇非偶函数"
            else:
                # 隐式方程对称性分析（关于x/y轴）
                x_sym = equation.subs({x: -x, y: y})  # x轴对称替换
                y_sym = equation.subs({x: x, y: -y})  # y轴对称替换
                sym_list = []
                if equation == x_sym: sym_list.append("关于y轴对称")
                if equation == y_sym: sym_list.append("关于x轴对称")
                parity = "；".join(sym_list) if sym_list else "无明显对称性"

            self.parity_label.setText(f"对称性: {parity}")

            # ------------------------- 计算零点/解集 -------------------------
            if is_explicit:
                zeros = solveset(expr, x, domain=S.Reals)  # 求解方程
                zero_text = self.set_to_readable(zeros, var='x')  # 转换为可读字符串
            else:
                try:
                    sol = solve(equation, y)  # 解隐式方程中的y
                    zero_text = "\n".join([f"y = {self.sympy_to_unicode(s)}" for s in sol])
                except:
                    zero_text = "复杂隐式解"

            self.zeros_label.setText(f"解集: {zero_text}")

            # ------------------------- 导数与单调性 -------------------------
            if is_explicit:
                derivative = diff(expr, x)  # 求导
                if is_linear:
                    mono_text = self.analyze_linear_monotonicity(derivative)
                else:
                    mono_text = self.analyze_general_monotonicity(derivative, x)
            else:
                try:
                    dy_dx = idiff(equation, y, x)  # 隐函数求导
                    mono_text = f"隐函数导数: dy/dx = {self.sympy_to_unicode(dy_dx)}"
                except:
                    mono_text = "隐函数导数复杂"

            self.monotonicity_label.setText(f"导数特性:\n{mono_text}")

            # ------------------------- 极值点分析 -------------------------
            if is_explicit and not is_linear:
                extrema_text = self.find_extrema(expr, x)
            else:
                extrema_text = "无"
            self.extrema_label.setText(f"极值点: {extrema_text}")

            # ------------------------- 渐近线分析 -------------------------
            if is_explicit:
                asymptotes = self.find_asymptotes(expr, x)
            else:
                asymptotes = "隐函数渐近线需具体分析"
            self.asymptotes_label.setText(f"渐近线: {asymptotes}")

            # ------------------------- 值域分析 -------------------------
            if is_explicit:
                range_text = self.estimate_range(expr, x)
            else:
                range_text = "隐函数值域需具体分析"
            self.range_label.setText(f"值域: {range_text}")

        except Exception as e:
            # 所有标签重置为"无"
            labels = [
                self.domain_label, self.range_label, self.parity_label,
                self.monotonicity_label, self.extrema_label, self.zeros_label,
                self.asymptotes_label
            ]
            for label in labels:
                label.setText(label.text().split(":")[0] + ": 无")
            print(f"数学属性计算错误: {e}")

    def estimate_range(self, expr, x):
        """估算值域（处理一次、二次函数及其他）"""
        from sympy import degree, N, expand
        try:
            expr_expanded = expand(expr)  # 展开表达式

            # 一次函数值域为全体实数
            if expr_expanded.is_polynomial(x) and degree(expr_expanded, x) == 1:
                return "(-∞, ∞)"

            # 二次函数通过顶点计算值域
            if expr_expanded.is_polynomial(x) and degree(expr_expanded, x) == 2:
                a = expr_expanded.coeff(x, 2)
                b = expr_expanded.coeff(x, 1)
                vertex_x = -b / (2 * a)
                vertex_y = expr_expanded.subs(x, vertex_x)
                if a > 0:
                    return f"[{N(vertex_y, 3):.2f}, ∞)"
                else:
                    return f"(-∞, {N(vertex_y, 3):.2f}]"

            # 数值估算（在大范围内采样）
            f = lambdify(x, expr, 'numpy')
            x_vals = np.linspace(-1e12, 1e12, 1000)  # 大范围采样
            y_vals = f(x_vals)
            valid = y_vals[np.isfinite(y_vals)]  # 过滤无效值

            if len(valid) == 0:
                return "无实数范围"
            return f"[{np.min(valid):.2f}, {np.max(valid):.2f}]"

        except Exception as e:
            print(f"[DEBUG] 值域计算错误: {e}")
            return "无法估算"

    def set_to_readable(self, s, var='x'):
        """将SymPy集合转换为可读字符串"""
        from sympy import FiniteSet, Interval, Union
        if isinstance(s, FiniteSet):  # 离散解集
            return ", ".join([f"{var} = {self.sympy_to_unicode(e)}" for e in s])
        return self.sympy_to_unicode(s)  # 其他类型直接转换

    def analyze_linear_monotonicity(self, derivative):
        """分析线性函数的单调性"""
        if derivative.is_positive:
            return "严格递增函数"
        elif derivative.is_negative:
            return "严格递减函数"
        return "常函数"

    def analyze_general_monotonicity(self, derivative, x):
        """分析一般函数的单调性（通过导数的符号）"""
        critical = solveset(derivative, x, S.Reals)  # 临界点
        inc = solveset(derivative > 0, x, S.Reals)  # 导数大于0的区间
        dec = solveset(derivative < 0, x, S.Reals)  # 导数小于0的区间

        parts = []
        if not inc.is_EmptySet:
            parts.append(f"增区间: {self.set_to_readable(inc)}")
        if not dec.is_EmptySet:
            parts.append(f"减区间: {self.set_to_readable(dec)}")
        return "\n".join(parts)  # 格式化为多行文本

    def find_extrema(self, expr, x):
        """寻找极值点（通过一阶和二阶导数）"""
        derivative = diff(expr, x)
        critical_points = solveset(derivative, x, S.Reals)  # 一阶导零点

        extrema = []
        for cp in critical_points:
            second_deriv = diff(derivative, x).subs(x, cp)  # 二阶导数
            if second_deriv.is_positive:
                extrema.append(f"极小值点: {self.sympy_to_unicode(cp)}")
            elif second_deriv.is_negative:
                extrema.append(f"极大值点: {self.sympy_to_unicode(cp)}")
        return ", ".join(extrema) if extrema else "无"

    def find_asymptotes(self, expr, x):
        """寻找渐近线（垂直、水平、斜渐近线）"""
        asymptotes = []

        # 垂直渐近线（奇点）
        singular = singularities(expr, x, S.Reals)
        if singular:
            asymptotes.extend([f"x={self.sympy_to_unicode(s)}" for s in singular])

        # 斜渐近线（计算斜率m和截距b）
        try:
            m = limit(expr / x, x, oo)
            if m.is_finite:
                b = limit(expr - m * x, x, oo)
                asymptotes.append(f"y={self.sympy_to_unicode(m)}x + {self.sympy_to_unicode(b)}")
        except:
            pass

        return "; ".join(asymptotes) if asymptotes else "无"

    def sympy_to_unicode(self, expr):
        """将SymPy表达式转换为Unicode字符串（美化输出）"""
        return pretty(expr, use_unicode=True)

    def select_color(self):
        """打开颜色对话框选择线条颜色"""
        color = QColorDialog.getColor()
        if color.isValid():
            self.color_button.setStyleSheet(f"background-color: {color.name()}")
            self.apply_changes_to_line()  # 立即应用颜色

    def apply_changes_to_line(self):
        """将属性面板的修改应用到选中线条"""
        if self.selected_line is None:
            return

        # 获取当前属性值
        color = self.color_button.styleSheet().split("background-color:")[1].split(";")[0].strip()
        style = self.line_style_combo.currentText()
        width = self.line_width_spinbox.value()
        alpha = self.alpha_slider.value() / 100.0  # 转换为0-1范围

        # 更新线条属性
        self.selected_line.set_color(color)
        self.selected_line.set_linestyle(style)
        self.selected_line.set_linewidth(width)
        self.selected_line.set_alpha(alpha)
        self.canvas.draw()

    def update_properties_panel(self, line):
        """更新属性面板的控件值（与选中线条的属性同步）"""
        try:
            color = line.get_color()
            style = line.get_linestyle()
            width = line.get_linewidth()
            alpha = line.get_alpha()

            self.color_button.setStyleSheet(f"background-color: {color}")
            self.line_style_combo.setCurrentText(style)
            self.line_width_spinbox.setValue(int(width))
            self.alpha_slider.setValue(int(alpha * 100))  # 转换为0-100范围
        except Exception as e:
            print("error")

    def toggle_left_panel(self):
        """切换左侧面板的折叠与展开"""
        if self.collapsed:
            self.splitter.setSizes([350, 850])  # 恢复分割器比例
            self.left_widget.show()
            self.collapse_button.setText("折叠")
            self.expand_button.setVisible(False)
        else:
            self.splitter.setSizes([0, 1200])  # 完全隐藏左侧
            self.left_widget.hide()
            self.collapse_button.setText("展开")
            self.expand_button.setVisible(True)
            self.expand_button.move(20, 20)

        self.collapsed = not self.collapsed


if __name__ == "__main__":
    app = QApplication(sys.argv)  # 创建应用实例
    main_window = MainWindow()  # 创建主窗口
    main_window.show()  # 显示窗口
    # for i in range(99):
    #     main_window.add_item()
    sys.exit(app.exec_())  # 进入事件循环
