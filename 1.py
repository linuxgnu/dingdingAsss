# -*- coding: UTF-8 -*-
import time
# 必备的注释文件
# import pywinctl as pygetwindow
from PIL import ImageGrab,Image
import time
import cv2
import numpy as np
# import pyautogui
# import easyocr
import  os
import pytesseract
import pyperclip
# from win10toast import ToastNotifier
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import ollama
import threading

# from AppKit import NSWorkspace, NSScreen
from Quartz import CGWindowListCopyWindowInfo, CGWindowListCreateImage, kCGWindowListOptionAll, kCGNullWindowID
from Quartz import CGImageDestinationCreateWithURL, CGImageDestinationAddImage, CGImageDestinationFinalize
from Quartz import kCGWindowImageBoundsIgnoreFraming, CGRectNull, CGImageGetWidth, CGImageGetHeight
from Quartz import kCGWindowListOptionIncludingWindow
from CoreServices import kUTTypePNG
from CoreFoundation import NSURL
import subprocess
import re
import imagehash

log_area = ''

def get_window_id():
    command = 'osascript -e \'tell app "System Events" to get id of window 1 of application "钉钉"\''
    result = subprocess.check_output(command, shell=True).decode('utf-8').strip()
    return result

def list_windows():
    """列出所有窗口及其 ID 和名称"""
    options = kCGWindowListOptionAll
    window_list = CGWindowListCopyWindowInfo(options, kCGNullWindowID)
    
    for window in window_list:
        window_id = window.get('kCGWindowNumber', 0)
        window_name = window.get('kCGWindowName', '')
        owner_name = window.get('kCGWindowOwnerName', '')
        is_onscreen = window.get('kCGWindowIsOnscreen', False)
        if window_name:  # 只打印有名称的窗口
            print(f"Window ID: {window_id}, Owner: {owner_name}, Name: {window_name}, Onscreen: {is_onscreen}")
    
    return window_list

def capture_window(window_id, output_path="~/Desktop/screenshot.png"):
    """捕获指定窗口 ID 的截图并保存为 PNG 文件"""
    image = CGWindowListCreateImage(
        CGRectNull,
        kCGWindowListOptionIncludingWindow,
        window_id,
        kCGWindowImageBoundsIgnoreFraming
    )
    
    if image is None:
        print(f"错误：无法捕获窗口 ID {window_id} 的内容，可能是最小化、隐藏或 ID 无效")
        return False
    
    width = CGImageGetWidth(image)
    height = CGImageGetHeight(image)
    # print(f"调试：图像对象已创建，大小 {width} x {height}")
    
    output_path = os.path.expanduser(output_path)
    # print(f"调试：目标保存路径为 {output_path}")
    
    output_dir = os.path.dirname(output_path)
    if not os.access(output_dir, os.W_OK):
        print(f"错误：路径 {output_dir} 不可写，请检查权限")
        return False
    
    url = NSURL.fileURLWithPath_(output_path)
    if url is None:
        print("错误：无法创建文件 URL")
        return False
    
    dest = CGImageDestinationCreateWithURL(url, kUTTypePNG, 1, None)
    if dest is None:
        print("错误：无法创建图像目标，可能与文件路径或格式有关")
        return False
    
    CGImageDestinationAddImage(dest, image, None)
    success = CGImageDestinationFinalize(dest)
    
    if success:
        print(f"成功：截图已保存到 {output_path}")
    else:
        print("错误：保存截图失败，可能是图像数据无效或权限不足")
    
    return success

def restore_and_capture(window_id, app_name, output_path="~/Downloads/screenshot.png"):
    """临时显示窗口并截图"""
    # 检查窗口是否存在
    script_check = f'''
    tell application "System Events"
        tell process "{app_name}"
            if exists window 1 then
                return true
            else
                return false
            end if
        end tell
    end tell
    '''
    result = subprocess.run(['osascript', '-e', script_check], capture_output=True, text=True)
    if result.stdout.strip() != "true":
        print(f"错误：应用程序 '{app_name}' 的窗口不存在或无法访问")
        return False
    
    # 显示窗口并截图
    script_restore = f'''
    tell application "System Events"
        tell process "{app_name}"
            set visible of window 1 to true
            delay 0.1
        end tell
    end tell
    '''
    subprocess.run(['osascript', '-e', script_restore], capture_output=True, text=True)
    success = capture_window(window_id, output_path)
    
    # 恢复窗口状态
    script_hide = f'''
    tell application "System Events"
        tell process "{app_name}"
            set visible of window 1 to false
        end tell
    end tell
    '''
    subprocess.run(['osascript', '-e', script_hide], capture_output=True, text=True)
    return success


def get_window_id(app_name):
    script_restore = f'''
        tell app "System Events" to get id of window 1 of application "{app_name}"
        '''
    result = subprocess.run(['osascript', '-e', script_restore], capture_output=True, text=True)
    result = int(result.stdout)
    return result

# def toastmsg(msg):
#
#     toaster = ToastNotifier()
#     toaster.show_toast("钉钉回复工具", msg, duration=10)

# 打开对话框
def openchat(xm,ym):
    # # 显示结果
    # cv2.imshow('Detected Red Points', image)
    # cv2.waitKey(0)
    # cv2.destroyAllWindows()
    # 要点击屏幕上的那个点
    # 移动鼠标到图标位置
    pyautogui.moveTo(xm, ym, duration=1)
    time.sleep(2)
    # 点击图标
    pyautogui.click(xm, ym)

#对比图片
def phash_compare(img1_path, img2_path, hash_size=8):
    img1 = Image.open(img1_path).convert('L')  # 转为灰度图
    img2 = Image.open(img2_path).convert('L')
    
    hash1 = imagehash.phash(img1, hash_size=hash_size)
    hash2 = imagehash.phash(img2, hash_size=hash_size)
    
    diff = hash1 - hash2  # 计算汉明距离
    threshold = 5  # 设置阈值，值越小越严格
    return diff <= threshold

# 识别对话框中的文字
def watchtext(imgurl):
    print('识别图片', imgurl)
    # 读取图片
    image = cv2.imread(imgurl)

    img = Image.open(imgurl)
    # 获取图片大小
    img_size = img.size
    # h = img_size[1] #图片高度
    # w = img_size[0] #图片宽度
    # 设置截取部分相对位置
    x = 0.20 * img_size[0]+300
    y = 0.1 * img_size[1]
    # y = 350
    w = 1 * img_size[0]-400
    h = 1* img_size[1]-800
    # 截取图片
    cropped = img.crop((x, y, x + w, y + h))  # (x1,y1,x2,y2)
    # 保存截图图片，命名为test.png
    text_img_url = '~/Downloads/test01.png'
    text_img_url = os.path.expanduser(text_img_url)
    # print('进入重复检查：',os.path.exists(text_img_url))
    #对图片进行对比
    if os.path.exists(text_img_url):
        new_text_img_url = '~/Downloads/test02.png'
        new_text_img_url = os.path.expanduser(new_text_img_url)
        cropped.save(new_text_img_url)
        # print('重复检查结果：', phash_compare(text_img_url, new_text_img_url))
        if phash_compare(text_img_url, new_text_img_url) == True:
            os.remove(new_text_img_url)
            return False
    cropped.save(text_img_url)

    #使用视觉大模型进行OCR
    # response = ollama.chat(
    #     model="llama3.2-vision:11b",
    #     messages=[
    #         {"role": "user", "content": "充当 OCR 助手。分析提供的图像并用中文：\n1. 尽可能准确地识别图像中所有可见文本。\n2. 保留文本的原始结构和格式。\n3.如果是多人对话，请根据不同的人员的发言来分别识别内容。\n3. 如果任何单词或短语不清楚，请在转录中用 [不清楚] 表示。\n仅提供转录，不提供任何其他评论，也不需要特殊的开头说明，直接给出结果。", "images": [text_img_url]}
    #     ],
    # )

    # print(response)
    # exit();

    

    # 图片预处理，例如灰度化、二值化等 使用google Tesseract 识别
    image = cv2.imread(text_img_url)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    # 使用pytesseract识别文字
    # pytesseract.pytesseract.tesseract_cmd = r'/usr/local/Cellar/tesseract/5.3.3/bin/tesseract'
    pytesseract.pytesseract.tesseract_cmd = r'tesseract'
    text = pytesseract.image_to_string(thresh, lang='chi_sim')
    # print(text,'12312312312312')
    text = text.replace(" ","")
    text_arr = text.split('\n')
    result = []
    for t in text_arr:
        # if re.findall('[0-9]{1,2}:[0-9]{1,2}$', t):
        #     print(t, len(re.findall('[0-9]{1,2}:[0-9]{1,2}$', t)))
        #     exit()
        if t != '' and t != '已读' and len(re.findall('[0-9]{1,2}:[0-9]{1,2}$', t)) == 0 and len(t) >= 10:
            # print('识别到：',t,"\n")
            result.append(t)
    # print(result)
    # exit()
    # # 另外一个工具
    # 设为中英文混合识别：ch_sim en
    # reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
    # 识别图片
    #
    # print(str(os.path) + '/' + imgurl)
    # result = reader.readtext(imgurl, detail=0)
    # for i in result:
    #     # 输出识别出的信息
    #     # 输出识别出的信息
    #     # print('输出识别出的信息')
    #     print(i, end='')
    #     做一下图片的裁剪再识别：ch_sim en
    
    # exit()
    # 设为中英文混合识别：ch_sim en
    # reader = easyocr.Reader(['ch_sim', 'en'], gpu=False, verbose=False)
    # # 路径改为用户需要识别的图片的路径
    # result = reader.readtext(text_img_url, detail=0)
    # for i in result:
    #     # 输出识别出的信息
    #     # 输出识别出的信息
    #     # print('输出识别出的信息')
    #     print(i, end='')
    # print(result)
    # exit()
    return result

# 截图保存
def getmscreen():
    windowsjiantou  = pygetwindow.getWindowsWithTitle('钉钉')
    windowsjiantou[0].show()
    w = windowsjiantou[0]
    w.activate()
    # 获取桌面窗口的坐标和尺寸
    left, top, width, height = w.left, w.top, w.width, w.height
    w.activate()
    w.show()
    # 将窗口最大化
    #w.maximize()
    # 下面的单位是5秒
    time.sleep(0.5)
    # print('运行到了这里')
    # 使用ImageGrab.grab()方法截取桌面
    screenshot = ImageGrab.grab(bbox=(left, top, left + width, top + height))
    # 获取当前时间的时间戳
    timestamp = time.time()
    # print("当前时间戳：", timestamp)
    imgurl =  str(timestamp)+'desktop_screenshot.png'
    # 保存截图
    # screenshot.save(imgurl)
    img = pyautogui.screenshot()
    img.save(str(timestamp)+'desktop_screenshot.png')
    return imgurl

#与大模型交互
def getchat(questiontext):
    print('当前对话最后识别内容：',questiontext)
    # 接入质谱AI的API
    # client = ZhipuAI(api_key=" . ")  # 请填写您自己的APIKey
    content = "根据对话内容\n\"\"\"\n" + questiontext + "\n\"\"\"\n思考最佳回答\n\n任何对话中你就是维刚，你需要直接做出回应\n\n回答过程要非常的职业化，这是办公场景。最重要的是不要讲出你自己的思考过程，不要重复问题，请你当做是你在跟对方交谈，直接给出回复内容即可。"
    # print(content)
    response=ollama.chat(
        model="qwq:latest",
        stream=False,
        messages=[
            {"role": "user","content": content}
        ],
        options={"temperature":0}
    )

    

    # response = client.chat.completions.create(
    #     model="glm-4",  # 填写需要调用的模型名称  OA表单中选不到项目的添加方法
    #     messages=[
    #         {"role": "user", "content": questiontext},
    #     ],
    #     tools=[
    #         {
    #             "type": "retrieval",
    #             "retrieval": {
    #                 "knowledge_id": " ",
    #                 "prompt_template": "从文档\n\"\"\"\n{{knowledge}}\n\"\"\"\n中找问题\n\"\"\"\n{{question}}\n\"\"\"\n的答案，找到答案就仅使用文档语句回答问题，找不到答案就用自身知识回答并且告诉用户该信息不是来自文档。\n不要复述问题，直接开始回答。"
    #             }
    #         }
    #     ],
    #     stream=True,
    # )
    resstr = ""
    # print(response)

    # print(response['message']['content'])
    content=response['message']['content']
    resstr=content.split("</think>")[-1]
    resstr=resstr.strip()
    # print(content)
    # exit()
    # for chunk in response:
    #     #print(chunk.choices[0].delta)
    #     resstr = resstr + str(chunk.choices[0].delta.content)
    #     # print(chunk.choices[0].delta.content)
    # print(resstr)
    # 做一个data，把数据返回去
    return resstr

#复制消息并发送
def pasttext(text):
    # windowsjiantou  = pygetwindow.getWindowsWithTitle('钉钉')
    # windowsjiantou[0].show()
    # w = windowsjiantou[0]
    # w.activate()
    # # 移动鼠标到目标位置（这里以屏幕坐标为例）
    # pyautogui.moveTo(600, 1000)
    # # 模拟鼠标点击
    # pyautogui.click()
    # 模拟键盘输入
    # pyautogui.typewrite('你好www', interval=0.2)
    # # 模拟按下Win键
    # pyautogui.press("win")
    # # 输入中文输入法的名称，例如“微软拼音输入法”
    # pyautogui.typewrite("微软拼音输入法")
    # # 模拟按下回车键
    # pyautogui.press("enter")
    # # 等待中文输入法启动
    # pyautogui.sleep(1)
    # # 输入中文字符
    # pyautogui.typewrite("你好，世界！")
    pyperclip.copy(text)
    # time.sleep(0.5)
    # pyautogui.hotkey('command', 'a')
    # pyautogui.hotkey('command', 'v')
    # time.sleep(0.5)
    # pyautogui.hotkey('enter')
    # pyperclip.paste()

#截图启动
def capture():
    # toastmsg('程序运行中')
    # 获取桌面窗口
    # desktop_window = pygetwindow.getDesktopWindow()
    # desktop_window = pygetwindow.getAllWindows()
    # desktop_window_title = pygetwindow.getAllTitles()
    # # for window in desktop_window_title:
    # #     print(window)
    # windowsjiantou  = pygetwindow.getWindowsWithTitle('钉钉')
    # print('捕获到钉钉')
    # windowsjiantou[0].show()
    # w = windowsjiantou[0]
    # w.activate()
    # # 获取桌面窗口的坐标和尺寸
    # left, top, width, height = w.left, w.top, w.width, w.height
    # w.activate()
    # w.show()
    # # 将窗口最大化
    # #w.maximize()
    # # 下面的单位是5秒
    # time.sleep(0.5)
    # # print('运行到了这里')
    # # 使用ImageGrab.grab()方法截取桌面
    # screenshot = ImageGrab.grab(bbox=(left, top, left + width, top + height))
    # # 获取当前时间的时间戳
    # timestamp = time.time()
    # # print("当前时间戳：", timestamp)
    # # 保存截图
    # imgs =str(timestamp)+'desktop_screenshot.png'
    # screenshot.save(imgs)
    # 读取图片上的红点
    # 识别图片
    # imgs =str(timestamp)+'desktop_screenshot.png'

    # window_list = list_windows()
    app_name = "钉钉"  # 替换为正确的应用程序名称，例如 "Safari"、"Finder"
    target_window_id = get_window_id(app_name) # 使用你的窗口 ID
    # log_area.insert(tk.END, "开始捕获", app_name,";\nID:", target_window_id)
    # print("开始捕获", app_name,";\nID:", target_window_id)
    imgs = "~/Downloads/screenshot.png"
    imgs = os.path.expanduser(imgs)
    restore_and_capture(target_window_id, app_name, imgs)
    time.sleep(0.5)
    # 读取图像
    image = cv2.imread(imgs)
    # 读取图像
    # 将图像从BGR转换为HSV颜色空间
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    # 定义红色的HSV范围
    lower_red1 = np.array([0, 120, 70])
    upper_red1 = np.array([10, 255, 255])
    lower_red2 = np.array([170, 120, 70])
    upper_red2 = np.array([180, 255, 255])
    # 创建掩码
    mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
    mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
    mask = cv2.bitwise_or(mask1, mask2)
    # 形态学操作以去除噪声
    kernel = np.ones((5, 5), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
    mask = cv2.dilate(mask, kernel, iterations=1)
    # 寻找轮廓  这里满足要求的轮廓已经放到这里数组里了
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    myusecolours = []
    # 绘制轮廓
    for contour in contours:
        # 计算轮廓的面积
        # 先留下面积大于100的轮廓
        area = cv2.contourArea(contour)
        if area > 50:  # 可以根据实际情况调整这个阈值
            # print('面积大于50')
            # 计算轮廓的周长
            perimeter = cv2.arcLength(contour, True)
            # 计算轮廓的近似形状
            # approxPolyDP 函数用于计算轮廓的近似形状
            # approxPolyDP
            approx = cv2.approxPolyDP(contour, 0.04 * perimeter, True)
            # 如果轮廓是圆形，那么近似形状的顶点数量应该接近于0
            # 但是这里我直接用半径来判断
            if len(approx) < 10:
                (x, y), radius = cv2.minEnclosingCircle(contour)
                center = (int(x), int(y))
                radius = int(radius)
                if radius > 5:  # 可以根据实际情况调整这个阈值  圆角值改小了一点
                    # 使用cv2.circle() 在原图上绘制筛选后的圆形轮廓。
                    # print('绘制了一个图形print')
                    cv2.circle(image, center, radius, (0, 255, 0), 2)
                    # 这里是通过考验的contour
                    # 获取contour 的坐标
                    # print(contour)
                    myusecolours.append(contour)
    # 显示结果
    # cv2.imshow('Contours', image)
    # cv2.waitKey(0)
    # cv2.destroyAllWindows()
    # print('----')
    myusecolours02 =myusecolours
    myusecolours02.reverse()
    # print(len(myusecolours02))
    # print(len(myusecolours02))
    if len(myusecolours02) == 0:
        return
    contoursmsg = myusecolours02[-1]
    # if len(myusecolours02) < 3:
    #     contoursmsg = myusecolours02[2]
    #
    #
    # # 获取第一条未读消息
    # if len(myusecolours02) < 2:
    #     contoursmsg = myusecolours02[1]
    #
    # if len(myusecolours02) < 1:
    #     contoursmsg = myusecolours02[0]
    # 获取坐标
    x, y, w, h = cv2.boundingRect(contoursmsg)
    # 打印边界框坐标
    # print(f"Bounding box coordinates: x={x}, y={y}, w={w}, h={h}")
    # 得到中心点的位置
    (xm, ym), radius = cv2.minEnclosingCircle(contoursmsg)
    # print(f"Bounding box coordinates: ----------------------------  x={xm}, y={ym}")
    # 打开对话框
    #openchat(xm,ym)
    # 截图
    # imgurl = getmscreen()
    imgurl = imgs
    # 识别对话框中的文字
    textcontent = watchtext(imgurl)
    if textcontent == False:
        print('图片无变化.本轮结束')
        os.remove(imgs)
        return False
    # print(textcontent)
    textcontent02 = ''
    text_res = []
    for item in textcontent:
        if item == '':
            continue
        # print(item+'\n')
        textcontent02= textcontent02+item+''
        text_res.append(item)
    # 获取最后一条消息
    # textcontent.reverse()
    # lasttext = textcontent[0]
    # print('最新的一条消息')
    # print(lasttext)
    # 调用API开始聊天--最后一条消息
    text_res.reverse()
    # print(text_res)
    # exit()
    textcontent01 = text_res[0]
    answer = getchat(textcontent01)
    # 调用API开始聊天--所有识别的内容
    # answer = getchat(textcontent02)
    # 将内容粘贴到钉钉窗口中
    print('AI的回应：', answer)
    pasttext(answer)

    os.remove(imgs)
    # os.remove(imgurl)
    # os.remove('test01.png')
    # toastmsg('程序运完毕')
    # print(desktop_window)
    # print(desktop_window_title)
    # # 获取桌面窗口的坐标和尺寸
    # left, top, width, height = desktop_window.left, desktop_window.top, desktop_window.width, desktop_window.height
    #
    # # 使用ImageGrab.grab()方法截取桌面
    # screenshot = ImageGrab.grab(bbox=(left, top, left + width, top + height))
    #
    # # 保存截图
    # screenshot.save('desktop_screenshot.png')

class MyException(Exception):
    pass

class MyThread(threading.Thread):
 
    def run(self):
        # Variable that stores the exception, if raised by someFunction
        self.exc = None           
        try:
            # self.someFunction()
            # log_area.insert(tk.END,'开干')
            capture()
        except BaseException as e:
            self.exc = e
       
    def join(self):
        threading.Thread.join(self)
        # Since join() returns in caller thread
        # we re-raise the caught exception
        # if any was caught
        if self.exc:
            raise self.exc

#启动一个新线程执行
def run_now():
    t = MyThread()
    t.start()
     
    # Exception handled in Caller thread
    try:
        t.join()
    except Exception as e:
        print("Exception Handled in Main, Details of the Exception:", str(e))

if __name__ == '__main__':
    # while True:
    #     run_now()
    #     time.sleep(5)
    run_now()
    # 先来屏幕截图
    #capture()
    # root = tk.Tk()
    # root.tk.call("tk", "scaling", 2.0)  # 调整缩放比例
    # root.geometry("400x500")
    # # 禁止用户调整窗口大小
    # root.resizable(False, False)
    # root.focus_force()
    # label = tk.Label(root, text=" ", font=("Microsoft YaHei", 16))
    # label.pack(pady=20)
    
    
    # label = tk.Label(root, text="点击 接管电脑 后，程序会识别未读消息并到知识库中进行检索填充回复。对信息修改勾，可以进行发送，或者设置自动发送",wraplength=300, font=("Microsoft YaHei", 16))
    # label.pack(pady=20)
    
    
    
    # button = tk.Button(root, text="接管电脑", command=run_now)
    # button.pack(pady=20)

    # log_area = ScrolledText(root, height=15, width=70, wrap=tk.WORD)
    # log_area.pack(fill="both", expand=True, pady=5)   
    # # log_area.insert(tk.END,'12312312312')
    # root.mainloop()

