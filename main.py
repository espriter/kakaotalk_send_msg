import sys
import pyautogui
import time
import pyperclip
import os
import random
import pandas as pd
import datetime


def send_msg(my_msg, repeat_number):
    for i in range(int(repeat_number)):
        time_wait = random.uniform(3, 5)
        print('Repeat Number : ', i + 1, end='')
        print(' // Time wait : ', time_wait)
        time.sleep(time_wait)
        pyautogui.keyDown('enter')
        pyperclip.copy(my_msg)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.keyDown('enter')
        pyautogui.keyDown('esc')
        pyautogui.keyDown('down')
        pyautogui.keyDown('down')


def filter_friend(filter_keyword, init_number):
    # 채팅 아이콘 클릭
    try:
        click_img(img_path + 'chat_icon.png')
        try:
            click_img(img_path + 'chat_icon2.png')
        except Exception as e :
            print('e ', e)
    except Exception as e :
        print('e ', e)
    # X 버튼이 존재한다면 클릭하여 내용 삭제
    try:
        click_img(img_path + 'x.png')
    except:
        pass
    time.sleep(1)
    # 돋보기 아이콘 오른쪽 클릭
    click_img_plus_x(img_path+'search_icon.png', 40)
    if filter_keyword == '':
        pyautogui.keyDown('esc')
    else:
        pyperclip.copy(filter_keyword)
    pyautogui.hotkey('ctrl', 'v')
    for i in range(int(init_number)-1):
        pyautogui.keyDown('down')
    time.sleep(2)

def return_status():
    # X 버튼이 존재한다면 클릭하여 내용 삭제
    try:
        click_img(img_path + 'x.png')
    except:
        pass
    time.sleep(1)
    # 친구 목록 아이콘 클릭해서 초기화
    try:
        click_img(img_path + 'person_icon.png')
        try:
            click_img(img_path + 'person_icon2.png')
        except Exception as e :
            print('e ', e)
    except Exception as e :
        print('e ', e)

def click_img(imagePath):
    location = pyautogui.locateCenterOnScreen(imagePath, confidence = conf)
    x, y = location
    pyautogui.click(x, y)


def click_img_plus_x(imagePath, pixel):
    pixel = 15
    location = pyautogui.locateCenterOnScreen(imagePath, confidence = conf)
    x, y = location
    pyautogui.click(x + pixel, y)


def doubleClickImg (imagePath):
    location = pyautogui.locateCenterOnScreen(imagePath, confidence = conf)
    x, y = location
    pyautogui.click(x, y, clicks=2)


def set_delay(delay_time):
    print(str(delay_time) + "초 후에 프로그램을 실행합니다.")
    for remaining in range(int(delay_time), 0, -1):
        sys.stdout.write("\r")
        sys.stdout.write("{:2d} seconds remaining.".format(remaining))
        sys.stdout.flush()
        time.sleep(1)
    sys.stdout.write("\r프로그램 실행!\n")


def logout():
    try:
        click_img(img_path + 'menu.png')
    except Exception as e:
        print('e ', e)
    try:
        click_img(img_path + 'logout.png')
    except Exception as e:
        print('e ', e)


def bye_msg():
    input('프로그램이 종료되었습니다.')


def set_import_msg():
    # 데이터 호출
    filter_keyword_data = pd.read_excel('./data/dataset.xlsx', sheet_name = 'FILTER_KEYWORD', header=0, engine='openpyxl')
    data_set_1 = pd.read_excel('./data/dataset.xlsx', sheet_name = 'CONTEXT_1', header=0, engine='openpyxl')
    data_set_2 = pd.read_excel('./data/dataset.xlsx', sheet_name = 'CONTEXT_2', header=0, engine='openpyxl')
    
    # 오늘 날짜 및 필터 키워드 확인
    filter_lst = []
    today_date = datetime.datetime.today().strftime("%Y-%m-%d")

    filter_keyword_1 = filter_keyword_data['CONTEXT_1'][0]
    filter_keyword_2 = filter_keyword_data['CONTEXT_2'][0]

    filter_lst.insert(0, filter_keyword_1)
    filter_lst.insert(1, filter_keyword_2)
    
    try:
        selected_contest_lst = []
        selected_context_1 = data_set_1[data_set_1['DATE'] == today_date]['CONTEXT'][0]
        selected_context_2 = data_set_2[data_set_2['DATE'] == today_date]['CONTEXT'][0]
        
        selected_contest_lst.insert(0, selected_context_1)
        selected_contest_lst.insert(1, selected_context_2)
    except Exception as e:
        print('해당 날짜가 없으므로 프로그램을 종료합니다.')
        print('e', e)
        quit()
    
    return filter_lst, selected_contest_lst


# config
img_path = os.path.dirname(os.path.realpath(__file__)) + '/img/'
conf = 0.90
pyautogui.PAUSE = 0.5

if __name__ == "__main__":
    print('Monitor size : ', end='')
    print(pyautogui.size())
    print(pyautogui.position())

    # 초기화 데이터 가져오기
    init_data = pd.read_excel('./data/dataset.xlsx', sheet_name = 'INIT_DATA', header=0, engine='openpyxl')
    init_number = int(init_data['INIT_NUM'][0]) # 몇번째 채팅방부터 시작할 지 / 0이면 첫번째 채팅방까지만 수행
    repeat_number = int(init_data['REPEAT_NUM'][0]) # 몇번째 채팅방까지 수행할 지 / 1이면 1개의 채팅방까지만 수행
    delay_time = int(init_data['DELAY_TIME'][0]) # 몇 초 뒤부터 시작할 지

    # 데이터 추출 및 지연 처리
    filter_lst, selected_contest_lst = set_import_msg()
    set_delay(delay_time)

    # 첫번째 공지 시작
    filter_friend(filter_lst[0], init_number)
    send_msg(selected_contest_lst[0], repeat_number)
    return_status()
    print('첫번째 공지가 종료되었습니다.')

    # 두번째 공지 시작
    filter_friend(filter_lst[1], init_number)
    send_msg(selected_contest_lst[1], repeat_number)
    return_status()
    print('두번째 공지가 종료되었습니다.')

    # 프로그램 종료
    bye_msg()
