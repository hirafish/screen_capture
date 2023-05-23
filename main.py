# 必要なモジュールのインポート
import pyocr
from PIL import Image, ImageEnhance
import os
import time
import datetime
import pyautogui
import pandas as pd
import ctypes
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from termcolor import colored
import numpy as np
import colorama

def main_capture() -> None:
    #colormaの初期化
    colorama.init()

    # 画面のキャプチャを行う
    capture()


def get_coordinate() -> tuple:
    '''
    スクリーンの座標を取得する
    '''
    clicked_once = 0
    if clicked_once == 0:
        try:
            print(colored("選択範囲の左上をクリックしてください", "white", attrs=['bold']))
            while True:
                # 左クリックが押されたら座標を取得
                if ctypes.windll.user32.GetAsyncKeyState(0x01) == 0x8000:
                    x1, y1 = pyautogui.position()
                    clicked_once = 1
                    break
        except KeyboardInterrupt:
            print(colored('----------終了----------', "red", attrs=["bold"]))

        time.sleep(1)

    if clicked_once == 1:
        try:
            print(colored("選択範囲の右下をクリックしてください", "white", attrs=['bold']))
            while True:
                if ctypes.windll.user32.GetAsyncKeyState(0x01) == 0x8000:
                    x2, y2 = pyautogui.position()
                    break

        except KeyboardInterrupt:
            print(colored('----------終了----------', "red", attrs=["bold"]))

    return x1, y1, x2, y2


def capture() -> None:
    name = input("PCのユーザ名を記入してください")
    x1, y1, x2, y2 = get_coordinate()

    # もし、左上と右下の座標が逆だった場合、エラーを出力して終了
    if x2 - x1 < 0 or y2 - y1 < 0:
        print(colored('Eroor : 「左上」と「右下」を選択してください', "red", attrs=["bold"]))
        print(colored('最初からやり直してください', "red", attrs=["bold"]))
        exit()

    # 初期のスクリーン状態として、画面のスクリーンショットを保存
    temp = pyautogui.screenshot(region = (x1, y1, x2 - x1, y2 - y1))
    # 保存先のディレクトリがなければ作成
    os.makedirs(f'C:/Users/{name}/Pictures/screen_capture', exist_ok=True)
    # 初期画像を保存
    temp.save(f"C:/Users/{name}/Pictures/screen_capture/0.png")
    print(colored("キャプチャ準備完了", "green", attrs=['bold']))
    img_cnt = 0
    capture_time_list = []
    capture_num_list = []

    try:
        while True:
            # 3秒ごとに画面を取得
            time.sleep(3)
            print(f"キャプチャ中...... ")

            # 画面を取得
            screenshot = pyautogui.screenshot(region = (x1, y1, x2 - x1, y2 - y1))
            # 3秒前の画像との差分を計算
            im_diff = np.array(temp).astype(int) - np.array(screenshot).astype(int)
            temp = pyautogui.screenshot(region = (x1, y1, x2 - x1, y2 - y1))
            area = (np.abs(im_diff) > 32).sum()

            # 差分が300以上なら画像を保存(300は自分で調整したもの)
            if area > 300:
                img_cnt += 1
                temp.save(f"C:/Users/{name}/Pictures/screen_capture/{img_cnt}.png")
                now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
                now_time = now.strftime('%H%M%S')
                # 保存した画像の名前と保存時間を表示
                print(colored(f"{now_time[:2]}時{now_time[2:4]}分{now_time[4:]}秒に{img_cnt}.pngとして保存しました", "green", attrs=["bold"]))
                # 保存した画像の名前と保存時間をcsvファイルに保存
                capture_time_list.append(f"{now_time[:2]}時{now_time[2:4]}分{now_time[4:]}秒")
                capture_num_list.append(f"{img_cnt}.png")
                dict = {'保存時間': capture_time_list, 'ファイル名': capture_num_list}
                df = pd.DataFrame(dict)
                df.to_csv(f"C:/Users/{name}/Pictures/screen_capture/保存時間.csv")

                # 保存した画像のOCRを行う
                img = Image.open(f"C://Users//{name}//Pictures//screen_capture//{img_cnt}.png")
                txt_pyocr = image_ocr(img)

                # OCRの結果内に特定の文字列があるかどうかを判定
                in_notify_word = check_text_in_notify_list(txt_pyocr)

                # 通知する文字列があればgooglespreadsheetに書き込み
                if in_notify_word:
                    write_google_spread_sheet(img_cnt)
                    print(colored("LINEに通知を発信しました", "green", attrs=["bold"]))

    except KeyboardInterrupt:
            print(colored('----------終了----------', "red", attrs=["bold"]))



def image_ocr(img) -> str:
    '''
    画像をOCRにかけ、結果として文字列を返す
    '''

    # 文字認識の精度を上げるための前処理
    img_g = img.convert('L') #gray変換
    enhancer= ImageEnhance.Contrast(img_g) #コントラストを上げる
    img_con = enhancer.enhance(2.0) #コントラストを上げる


    # OCRに用いるソフトウェアのPATHを指定
    TESSERACT_PATH = 'C:\\Program Files\\Tesseract-OCR' #インストールしたTesseract-OCRのpath
    TESSDATA_PATH = 'C:\\Program Files\\Tesseract-OCR\\tessdata' #tessdataのpath
    os.environ["PATH"] += os.pathsep + TESSERACT_PATH
    os.environ["TESSDATA_PREFIX"] = TESSDATA_PATH
    # OCRエンジン取得
    tools = pyocr.get_available_tools()
    tool = tools[0]
    # OCRの設定
    builder = pyocr.builders.TextBuilder(tesseract_layout=6)


    # OCRの実行->txt_pyocrに保存
    txt_pyocr = tool.image_to_string(img_con , lang='jpn', builder=builder)

    return txt_pyocr


def check_text_in_notify_list(txt_pyocr: str) -> bool:
    '''
    OCRの結果、特定の文字列が含まれていた場合に、trueを返す
    '''

    in_notify_word = False

    # この文字列が存在したら通知する
    check_list = ["重要", "課題", "期限", "宿題", "締め切り", "締切", "期間", "提出", "練習"]
    for word in check_list:
        if word in txt_pyocr:
            in_notify_word = True
            break

    return in_notify_word


def write_google_spread_sheet(img_cnt: int) -> None:
    '''
    OCRをした結果、特定の文字列が含まれていた場合に、spreadsheetを更新する。
    '''
    # Googleスプレッドシートの認証
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    # ダウンロードしたjsonファイルを同じフォルダに格納して指定
    json_keyfile_name = ""
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_name, scope)
    # 認証情報を使ってスプレッドシートの操作権を取得
    gc = gspread.authorize(credentials)

    # 共有設定したスプレッドシートのキーを指定
    SPREADSHEET_KEY = ''
    #共有設定したスプレッドシートのシート1を開く
    worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1
    #A1セルをimg_cntの値に更新する
    worksheet.update_cell(1, 1, img_cnt)

    return


main_capture()