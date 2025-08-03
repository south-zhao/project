import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import random
import os
import sys
import pandas as pd
import pygame
from datetime import datetime


class StudentPicker:
    def __init__(self, root):
        self.students = []
        self.picked_students = []
        self.root = root
        self.root.title("班级随机抽签系统")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f7fa")
        self.is_rolling = False

        # 初始化pygame混音器
        pygame.mixer.init()

        # 加载音效
        self.sounds_loaded = False
        self.load_sounds()

        # 创建UI
        self.create_widgets()

        # 加载学生名单
        self.student_file = "students.xlsx"
        self.load_students()

        # 设置定时器ID
        self.after_id = None
        self.roll_speed = 100  # 滚动速度(毫秒)

        # 播放背景音乐
        self.play_background_music()

        # 设置关闭事件处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def resource_path(self, relative_path):
        """获取资源的绝对路径，支持PyInstaller打包"""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def load_sounds(self):
        """尝试加载音效文件"""
        try:
            # 创建音效目录
            sound_dir = "sounds"
            os.makedirs(sound_dir, exist_ok=True)

            # 背景音乐
            bg_music_path = os.path.join(sound_dir, "background_music.mp3")
            self.background_music = pygame.mixer.Sound(bg_music_path)

            # 滚动音效
            rolling_sound_path = os.path.join(sound_dir, "rolling_sound.mp3")
            self.rolling_sound = pygame.mixer.Sound(rolling_sound_path)

            # 选中音效
            select_sound_path = os.path.join(sound_dir, "select_sound.mp3")
            self.select_sound = pygame.mixer.Sound(select_sound_path)

            self.sounds_loaded = True
        except Exception as e:
            print(f"音效加载失败: {e}")
            self.sounds_loaded = False


    def play_background_music(self):
        """播放背景音乐"""
        if self.sounds_loaded:
            try:
                self.background_music.set_volume(0.3)
                self.background_music.play(loops=-1)  # 循环播放
            except:
                self.sounds_loaded = False

    def load_students(self):
        """从Excel文件加载学生名单"""
        self.initial_students = []
        self.students = []
        self.picked_students = []

        try:
            if os.path.exists(self.student_file):
                # 读取Excel文件
                df = pd.read_excel(self.student_file)
                if '姓名' in df.columns:
                    self.initial_students = df['姓名'].tolist()
                else:
                    # 如果没有姓名列，使用第一列
                    self.initial_students = df.iloc[:, 0].tolist()

                # 确保名单不为空
                if not self.initial_students:
                    self.initial_students = []

                self.students = self.initial_students.copy()
                self.update_status(f"成功加载 {len(self.students)} 名学生")
            else:
                # 创建默认名单
                self.initial_students = []
                self.students = self.initial_students.copy()
                self.save_students()
                self.update_status(f"创建默认名单，共 {len(self.students)} 名学生")
        except Exception as e:
            self.initial_students = []
            self.students = self.initial_students.copy()
            self.update_status(f"加载名单错误: {e}, 使用默认名单")

        # 更新名单显示
        self.update_student_lists()

    def save_students(self):
        """保存学生名单到Excel"""
        df = pd.DataFrame({"姓名": self.initial_students})
        df.to_excel(self.student_file, index=False)


    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=10)

        title_label = ttk.Label(
            title_frame,
            text="班级随机抽签系统",
            font=("微软雅黑", 24, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack(side=tk.LEFT)

        # 控制按钮
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=10)

        self.start_btn = ttk.Button(
            control_frame,
            text="开始抽签",
            command=self.toggle_roll,
            width=15,
        )
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.reset_btn = ttk.Button(
            control_frame,
            text="重新开始",
            command=self.reset,
            width=15
        )
        self.reset_btn.pack(side=tk.LEFT, padx=5)

        self.import_btn = ttk.Button(
            control_frame,
            text="导入名单",
            command=self.import_students,
            width=15
        )
        self.import_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(
            control_frame,
            text="导出结果",
            command=self.export_results,
            width=15
        )
        self.export_btn.pack(side=tk.LEFT, padx=5)

        # 显示区域
        display_frame = ttk.LabelFrame(
            main_frame,
            text="名字显示区",
            padding=20
        )
        display_frame.pack(fill=tk.BOTH, pady=10, expand=True)

        self.display_var = tk.StringVar()
        self.display_var.set("准备开始")
        self.display_label = ttk.Label(
            display_frame,
            textvariable=self.display_var,
            font=("微软雅黑", 36, "bold"),
            foreground="#3498db",
            anchor=tk.CENTER
        )
        self.display_label.pack(fill=tk.BOTH, expand=True)

        # 速度控制
        speed_frame = ttk.Frame(display_frame)
        speed_frame.pack(fill=tk.X, pady=10)

        ttk.Label(speed_frame, text="滚动速度:").pack(side=tk.LEFT, padx=5)
        self.speed_var = tk.IntVar(value=100)
        speed_scale = ttk.Scale(
            speed_frame,
            from_=50,
            to=300,
            variable=self.speed_var,
            command=self.update_speed
        )
        speed_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.speed_label = ttk.Label(speed_frame)
        self.speed_label.pack(side=tk.LEFT, padx=5)

        # 名单显示区
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 未选名单
        available_frame = ttk.LabelFrame(
            list_frame,
            text="待选名单",
            padding=10
        )
        available_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self.available_list = tk.Listbox(
            available_frame,
            font=("微软雅黑", 12),
            selectmode=tk.SINGLE,
            exportselection=False
        )
        self.available_list.pack(fill=tk.BOTH, expand=True)

        # 已选名单
        selected_frame = ttk.LabelFrame(
            list_frame,
            text="已选名单",
            padding=10
        )
        selected_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self.selected_list = tk.Listbox(
            selected_frame,
            font=("微软雅黑", 12),
            selectmode=tk.SINGLE,
            exportselection=False
        )
        self.selected_list.pack(fill=tk.BOTH, expand=True)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding=(10, 5)
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 配置样式
        self.style = ttk.Style()
        self.style.configure("Accent.TButton", background="#3498db", foreground="white")
        self.style.map("Accent.TButton",
                       background=[('active', '#2980b9'), ('pressed', '#1c638e')])

    def update_speed(self, value):
        """更新滚动速度"""
        self.roll_speed = int(float(value))

    def update_status(self, message):
        """更新状态栏"""
        self.status_var.set(message)

    def update_student_lists(self):
        """更新学生名单显示"""
        # 更新待选名单
        self.available_list.delete(0, tk.END)
        for student in self.students:
            self.available_list.insert(tk.END, student)

        # 更新已选名单
        self.selected_list.delete(0, tk.END)
        for student in self.picked_students:
            self.selected_list.insert(tk.END, student)

    def toggle_roll(self):
        """开始/停止滚动"""
        if not self.is_rolling:
            # 开始滚动
            if len(self.students) == 0:
                self.update_status("名单已空，无法点名!")
                return

            self.is_rolling = True
            self.start_btn.config(text="停止点名")
            self.update_status("名字滚动中...")

            # 播放滚动音效
            if self.sounds_loaded:
                try:
                    self.rolling_sound.set_volume(0.5)
                    self.rolling_sound.play(loops=-1)
                except:
                    self.sounds_loaded = False

            self.start_rolling()
        else:
            # 停止滚动
            self.is_rolling = False
            self.start_btn.config(text="开始点名")
            self.update_status("已停止滚动")

            # 停止滚动音效
            if self.sounds_loaded:
                try:
                    self.rolling_sound.stop()
                except:
                    pass

            # 播放选中音效
            if self.sounds_loaded:
                try:
                    self.select_sound.set_volume(0.7)
                    self.select_sound.play()
                except:
                    self.sounds_loaded = False

            self.pick_student()

    def start_rolling(self):
        """开始名字滚动"""
        self.display_label.config(foreground="#3498db")
        if self.is_rolling:
            # 随机选择一个名字显示
            current_name = random.choice(self.students)
            self.display_var.set(current_name)

            # 设置下一次滚动
            self.after_id = self.root.after(self.roll_speed, self.start_rolling)

    def pick_student(self):
        """选取当前显示的学生"""
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None

        if len(self.students) > 0:
            # 从当前显示的名字中确定选中者
            picked = self.display_var.get()
            if picked not in self.students:
                picked = random.choice(self.students)

            # 记录并显示结果
            self.display_var.set(f"★ {picked} ★")
            self.display_label.config(foreground="#e74c3c")

            # 从列表中移除
            if picked in self.students:
                self.students.remove(picked)
                self.picked_students.append(picked)

            # 更新名单显示
            self.update_student_lists()

            # 更新状态
            self.update_status(f"已选中: {picked} | 剩余学生: {len(self.students)}人")

    def reset(self):
        """重置系统"""
        # 停止滚动
        self.is_rolling = False
        self.start_btn.config(text="开始点名")

        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None

        # 停止滚动音效
        if self.sounds_loaded:
            try:
                self.rolling_sound.stop()
            except:
                pass

        # 重置名单
        self.students = self.initial_students.copy()
        self.picked_students = []

        # 重置显示
        self.display_var.set("准备开始")
        self.display_label.config(foreground="#3498db")

        # 更新名单显示
        self.update_student_lists()

        # 更新状态
        self.update_status(f"已重置系统 | 总学生数: {len(self.students)}人")

    def import_students(self):
        """导入学生名单"""
        file_path = filedialog.askopenfilename(
            title="选择学生名单",
            filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
        )

        if file_path:
            try:
                # 读取Excel文件
                df = pd.read_excel(file_path)
                if '姓名' in df.columns:
                    self.initial_students = df['姓名'].tolist()
                else:
                    # 如果没有姓名列，使用第一列
                    self.initial_students = df.iloc[:, 0].tolist()

                # 确保名单不为空
                if not self.initial_students:
                    messagebox.showwarning("导入失败", "名单为空!")
                    return

                # 更新名单
                self.students = self.initial_students.copy()
                self.picked_students = []
                self.student_file = file_path

                # 保存名单
                self.save_students()

                # 更新显示
                self.reset()
                self.update_status(f"成功导入 {len(self.students)} 名学生")
            except Exception as e:
                messagebox.showerror("导入错误", f"无法导入名单: {e}")

    def export_results(self):
        """导出点名结果"""
        if not self.picked_students:
            messagebox.showinfo("导出结果", "没有已点名的学生!")
            return

        # 生成文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"点名结果_{timestamp}.xlsx"

        file_path = filedialog.asksaveasfilename(
            title="保存点名结果",
            initialfile=default_filename,
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )

        if file_path:

            # 创建DataFrame
            result_data = {
                "已点名学生": self.picked_students,
                "时间": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")] * len(self.picked_students)
            }

            # 添加未点名学生
            if self.students:
                result_data["未点名学生"] = self.students + [""] * (len(self.picked_students) - len(self.students))

            df = pd.DataFrame(result_data)

            # 保存到Excel
            df.to_excel(file_path, index=False)
            self.update_status(f"成功导出点名结果到: {os.path.basename(file_path)}")
            messagebox.showinfo("导出成功", f"点名结果已保存到:\n{file_path}")


    def on_closing(self):
        """关闭窗口时的处理"""
        # 停止所有音效
        if self.sounds_loaded:
            try:
                self.background_music.stop()
                self.rolling_sound.stop()
            except:
                pass

        # 保存名单
        try:
            self.save_students()
        except:
            pass

        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()

    app = StudentPicker(root)
    root.mainloop()