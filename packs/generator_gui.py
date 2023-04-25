from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from ttkthemes import ThemedStyle
from PIL import Image, ImageTk
from packs.generator_execute import GeneratorExecute


class GeneratorGui:
    def __init__(self, master):
        self.master = master
        self.sources = []
        self.templates = []
        self.rules = []
        self.outputs = []

        # 设置主题样式
        self.app_style = ThemedStyle(self.master)
        self.app_style.set_theme("arc")

        # 创建背景图片
        img = Image.open("pic/background.jpeg")
        img = img.resize((600, 350), Image.ANTIALIAS)
        self.background = ImageTk.PhotoImage(img)
        self.background_label = ttk.Label(self.master, image=self.background)
        self.background_label.place(x=0, y=0)

        # 定义新的 Style
        self.common_style = ttk.Style()
        self.common_style.configure(
            "Common.TButton",
            borderwidth=0,
            relief="flat",
            background="white",
            foreground="black",
            font=("Segoe UI", 12, "bold"),
            padding=10,
            bordercolor="black",
            borderradius=20,
        )
        self.common_style.configure(
            "Common.TLabel",
            borderwidth=0,
            relief="flat",
            background="white",
            foreground="black",
            font=("Segoe UI", 12, "bold"),
            padding=10,
            bordercolor="black",
            borderradius=20,
        )

        self.source_button = ttk.Button(
            self.master,
            text="选择源文件",
            command=self.choose_source,
            style=("Rounded.TButton", "Common.TButton"),
            width=12,
        )
        self.template_button = ttk.Button(
            self.master,
            text="选择模版文件",
            command=self.choose_template,
            style=("Rounded.TButton", "Common.TButton"),
            width=12,
        )
        self.rule_button = ttk.Button(
            self.master,
            text="选择规则文件",
            command=self.choose_rule,
            style=("Rounded.TButton", "Common.TButton"),
            width=12,
        )
        self.output_button = ttk.Button(
            self.master,
            text="选择输出文件夹",
            command=self.choose_output,
            style=("Rounded.TButton", "Common.TButton"),
            width=12,
        )

        # 选择的地址
        self.source_label = ttk.Label(
            self.master, text="地址：未选择", style=("Rounded.TLabel", "Common.TLabel")
        )
        self.template_label = ttk.Label(
            self.master, text="地址：未选择", style=("Rounded.TLabel", "Common.TLabel")
        )
        self.rule_label = ttk.Label(
            self.master, text="地址：未选择", style=("Rounded.TLabel", "Common.TLabel")
        )
        self.output_label = ttk.Label(
            self.master, text="地址：未选择", style=("Rounded.TLabel", "Common.TLabel")
        )
        # 下方的两个按钮
        self.execute_button = ttk.Button(
            self.master,
            text="执行",
            command=self.execute,
            style=("Rounded.TButton", "Common.TButton"),
            width=12,
        )
        self.quit_button = ttk.Button(
            self.master,
            text="退出",
            command=self.master.quit,
            style=("Rounded.TButton", "Common.TButton"),
            width=12,
        )

        # 布局按钮和标签
        self.source_button.place(x=50, y=50)
        self.template_button.place(x=50, y=100)
        self.rule_button.place(x=50, y=150)
        self.output_button.place(x=50, y=200)
        self.execute_button.place(x=50, y=280)
        self.quit_button.place(x=450, y=280)

        self.source_label.place(x=200, y=55)
        self.template_label.place(x=200, y=105)
        self.rule_label.place(x=200, y=155)
        self.output_label.place(x=200, y=205)

    def choose_source(self):
        # 打开选择源文件的对话框
        source_file = filedialog.askopenfilename()
        if source_file:
            self.sources.append(source_file)
            self.source_label.configure(text=f"源文件：{source_file}")

    def choose_template(self):
        # 打开选择模版文件的对话框
        template_file = filedialog.askopenfilename()
        if template_file:
            self.templates.append(template_file)
            self.template_label.configure(text=f"模版文件：{template_file}")

    def choose_rule(self):
        # 打开选择规则文件的对话框
        rule_file = filedialog.askopenfilename()
        if rule_file:
            self.rules.append(rule_file)
            self.rule_label.configure(text=f"规则文件：{rule_file}")

    def choose_output(self):
        # 打开选择输出文件夹的对话框
        output_folder = filedialog.askdirectory()
        if output_folder:
            self.outputs.append(output_folder)
            self.output_label.configure(text=f"输出文件夹：{output_folder}")

    def execute(self):
        # 检查是否已经选择了所有必要的文件和文件夹
        if not self.sources or not self.templates or not self.rules or not self.outputs:
            messagebox.showwarning("提示", "请选择所有必要的文件和文件夹")
            return

        # 获取所需参数
        source_file = self.sources[-1]
        template_file = self.templates[-1]
        rule_file = self.rules[-1]
        output_folder = self.outputs[-1]

        # 执行excel_transform模块中的相关函数
        # 在此处传递所需参数
        eg = GeneratorExecute(source_file, template_file,
                              rule_file, output_folder)
        eg.execute()

        # 显示执行完成的信息
        messagebox.showinfo("提示", "执行完成")

    def get_sources(self):
        return self.sources

    def get_templates(self):
        return self.templates

    def get_rules(self):
        return self.rules

    def get_outputs(self):
        return self.outputs
