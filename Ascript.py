import win32gui
import win32ui
import win32con
import ctypes
import numpy as np
import cv2
import traceback

class AScript:
    def find_hwnd_information(self, window_title):
        """
        根据窗口标题获取窗口句柄及信息
        """
        hwnd = win32gui.FindWindow(None, window_title)

        if hwnd:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            print("*******************************")
            print(f"窗口句柄: {hwnd}")
            print("窗口位置：({}, {})".format(left, top))
            print("窗口大小：{} x {}".format(right - left, bottom - top))
            print("*******************************")
            return hwnd
        else:
            print("未找到窗口")
            return None

    def capture_window_background(self, hwnd, save_path=None):
        """
        后台捕获指定窗口内容并返回 OpenCV 格式的图像
        :param hwnd: 窗口句柄
        :param save_path: 保存路径，若为 None，则不保存
        :return: OpenCV 格式的图像
        """
        try:
            # 获取窗口大小
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]

            # 验证窗口尺寸
            if width <= 0 or height <= 0:
                print("窗口尺寸无效，无法截图。")
                return None

            # 创建兼容的设备上下文和位图
            hwnd_dc = win32gui.GetWindowDC(hwnd)
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()
            save_bitmap = win32ui.CreateBitmap()
            save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
            save_dc.SelectObject(save_bitmap)

            # 使用 ctypes 调用 PrintWindow
            PW_RENDERFULLCONTENT = 2
            result = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), PW_RENDERFULLCONTENT)

            if not result:
                print("PrintWindow 捕获失败，可能是目标窗口不支持。")
                return None

            # 获取位图数据
            bmp_info = save_bitmap.GetInfo()
            bmp_data = save_bitmap.GetBitmapBits(True)

            # 转换为 OpenCV 图像
            img = np.frombuffer(bmp_data, dtype='uint8')
            img = img.reshape((bmp_info['bmHeight'], bmp_info['bmWidth'], 4))

            # 转为 BGR 格式
            img_bgr = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

            # 保存到文件（可选）
            if save_path:
                cv2.imwrite(save_path, img_bgr)
                print(f"截图已保存到: {save_path}")

            return img_bgr

        except Exception as e:
            print(f"捕获窗口失败: {e}")
            traceback.print_exc()
            return None

        finally:
            # 释放资源
            if save_dc:
                save_dc.DeleteDC()
            if mfc_dc:
                mfc_dc.DeleteDC()
            if hwnd_dc:
                win32gui.ReleaseDC(hwnd, hwnd_dc)
            if save_bitmap:
                win32gui.DeleteObject(save_bitmap.GetHandle())


    def check_screenshot(self, hwnd):
        """
        检查窗口内容是否与截图匹配，并标记匹配区域
        """
        # 尝试加载截图
        screenshot_path = "screenshot.png"
        screenshot = cv2.imread(screenshot_path)

        if screenshot is None:
            print("未检测到截图文件：screenshot.png")
            return

        print("检测到截图文件：screenshot.png")

        # 获取窗口内容
        img = self.capture_window(hwnd)

        if img is None:
            print("无法捕获窗口内容")
            return

        # 转为灰度图像
        gray_window = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray_screenshot = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

        # 使用模板匹配
        result = cv2.matchTemplate(gray_window, gray_screenshot, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)

        print(f"模板匹配相似度: {max_val:.2f}")
        if max_val > 0.8:  # 设置匹配阈值
            print("窗口内容与截图匹配！")
            # 标记匹配区域
            cv2.rectangle(img, max_loc,
                        (max_loc[0] + gray_screenshot.shape[1], max_loc[1] + gray_screenshot.shape[0]),
                        (0, 255, 0), 2)

            # 显示标记结果
            cv2.imshow("Matched Result", img)
            cv2.waitKey(0)
            cv2.destroyAllWindows()
        else:
            print("窗口内容与截图不匹配！")
        









