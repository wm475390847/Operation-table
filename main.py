import tkinter

from packs.gui_generator import FileSelector
from packs.execute_generator import ExecuteGenerator


if __name__ == "__main__":
    root = tkinter.Tk()
    root.geometry("600x350")
    root.title("操作表生成器")

    fs = FileSelector(root)
    root.mainloop()

    # 通过FileSelector实例获取选择的文件列表
    sources = fs.get_sources()
    templates = fs.get_templates()
    rules = fs.get_rules()
    outputs = fs.get_outputs()

    print(f"源文件：{sources}")
    print(f"模版文件：{templates}")
    print(f"规则文件：{rules}")
    print(f"输出文件夹：{outputs}")
