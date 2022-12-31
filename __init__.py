import pyautogui  # 键鼠操作自动化库
import time  # 处理时间的标准库
import xlrd  # 读取excel表格的第三方库
import pyperclip  # 复制粘贴库


# 定义鼠标事件

# pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159
# clicktimes:点击次数 lorR img为图片名称 reTry为mouseClick的执行次数
def mouseClick(clickTimes, lOrR, img, reTry):  # 自定义鼠标点击图标的方法
    if reTry == 1:
        while True:
            # 获取图片的中心坐标                             #confidence的作用目前不知道，没有confidence也可以运行
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)  # 将图片的中心位置传给location
            if location is not None:  # interval允许的左右误差,duration允许的上下误差#按下的鼠标按键
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2,
                                button=lOrR)  # 点击传入的参数
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)


# 数据检查
# cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
# ctype     空：0
#           字符串：1
#           数字：2
#           日期：3
#           布尔：4
#           error：5
def dataCheck(sheet1):
    checkCmd = True
    # 行数检查
    if sheet1.nrows < 2:  # sheet.nrows;获取sheet表的有效行数
        print("没数据啊哥")
        checkCmd = False
    # 每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        # 输入数据应为数字
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
                                  and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0):
            print('第', i + 1, "行,第1列数据有毛病")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:  # 输入数据应为图片的名字,即字符串类型
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                checkCmd = False
        i += 1
    return checkCmd


# 任务
def mainWork(img):
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "left", img, reTry)
            print("单击左键", img)
        # 2代表双击左键
        elif cmdType.value == 2.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry)
            print("双击左键", img)
        # 3代表右键
        elif cmdType.value == 3.0:
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry)
            print("右键", img)
            # 4代表输入
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)  # 复制内容到剪切板
            pyautogui.hotkey('ctrl', 'v')  # 粘贴内容
            # hptkey的作用类似于payutogui.keyDown()和payutogui.keyUp()
            # 先从左向右执行paytogui.keyDown(),然后从右向左执行paytogui.keyUp()
            time.sleep(0.5)
            print("输入:", inputValue)
            # 5代表等待
        elif cmdType.value == 5.0:
            # 取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待", waitTime, "秒")
        # 6代表滚轮
        elif cmdType.value == 6.0:
            # 取图片名称
            scrooltimes = 0  # 滚动次数
            times = int(sheet1.row(i)[2].value)
            while scrooltimes < times:
                scrooltimes += 1
                scroll = sheet1.row(i)[1].value
                pyautogui.scroll(int(scroll))
                print("滚轮滑动", int(scroll), "距离")
        i += 1


if __name__ == '__main__':
    file = 'cmd.xls'
    # 打开文件
    wb = xlrd.open_workbook(filename=file)  # 打开excel文件，将文件地址传给wb
    # 通过索引获取表格sheet页
    # print(wb.nsheets)获取excel中表的页数
    # print(wb.sheets())获取excel中sheet表对象
    # print(wb.sheet_names())
    # print(wb.sheet_by_index(1))#按索引获取sheet对象
    # print(wb.sheet_by_name('Sheet1'))按Sheet表名获取sheet对象
    sheet1 = wb.sheet_by_index(1)  # wb.sheet_by_index(0)对应excel中的第一页的地址。wb.sheet_by_index(1)对应excel中的第二页
    print('欢迎牛马牌自动小程序~')
    # 数据检查
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        key = input('选择功能: 1.做一次 2.循环到死 \n')
        if key == '1':
            # 循环拿出每一行指令
            mainWork(sheet1)
        elif key == '2':
            while True:
                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")
    else:
        print('输入有误或者已经退出!')
