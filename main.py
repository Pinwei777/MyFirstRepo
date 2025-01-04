
# 导入模块
from Ascript import Ascript

# 使用示例
if __name__ == "__main__":

    script = Ascript()

    # 窗口标题
    window_title = "Albion Online Client"

    # 查找窗口并获取句柄
    hwnd = script.find_hwnd_information(window_title)

    if hwnd:
        # 创建子窗口显示内容
        script.display_window(hwnd)
    else:
        print("未找到指定窗口。")