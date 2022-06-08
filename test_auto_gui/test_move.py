import pyautogui
import win32gui
import ctypes
import pyperclip
import time

# クロームのアイコン画像
# 注意：本プログラムを使用する際は、自身の環境でクロームと、クロームの画像を用意する必要があります。
chrome_icon = r".\img\chrome_icon.png"


# クローム起動
def start_chrome():
    # 次の操作までのインターバル
    pyautogui.PAUSE = 1.7
    
    # フェールセーフ機能有効
    # 動作確認中、マウスを左上に移動することで「pyautogui」の処理が停止する
    pyautogui.FAILSAFE = True
    
    # ポップアップ
    pyautogui.alert(text="クロームがメインのディスプレイに表示されていることを確認してください",button="OK")

    # 指定の画像までカーソル移動
    move_cursor(chrome_icon)
    
    pyautogui.doubleClick()
    
    # クロームの起動待機時間
    time.sleep(3)
    
    # 日本語文字入力に対応していないため、ペーストしたい文字を事前に準備
    pyperclip.copy("ダブルクリックで起動しました")
    
    # ペーストショートカットキー
    pyautogui.hotkey("ctrl","v")
    
    # Excelで新規ファイル作成したら、デフォルトの名前は「Book1」になる
    # そのウィンドウがある状態で下記を実行すると、「Book1」はアクティブになる
    # active_window("Book1")

    
# メインのディスプレイ内でしか画像認識不可
# ディスプレイ上に表示させていないと検知できない(重なっていると検知不可)
# 画像認識には、認識したい画像を用意する必要がある。
# その画像は、Win標準機能の切り取り機能でないと精度が保証できない


    
# 対象の画像までカーソルを動かす
def move_cursor(img_path:str):
    
    # 対象の画像を認識
    localtion = pyautogui.locateCenterOnScreen(img_path,confidence=0.8,grayscale=True) # 見つからない場合Noneが変える
    
    # 認識の有無確認
    if localtion is not None:
        # 認識がされた場合
        
        # 画像までカーソルを移動
        pyautogui.moveTo(localtion) # 見つかった場合はlocationにX,Y座標が入る
    else:
        # 認識されなかった場合
        print("見つかりませんでした")
        exit()


# 指定したウィンドウをアクティブにする
def active_window(target_title:str):

    def forground( hwnd, title):
        name = win32gui.GetWindowText(hwnd)
        if name.find(title) >= 0:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd,1) # SW_SHOWNORMAL
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            return False

    try:
        win32gui.EnumWindows(forground, target_title)
    except:
        pass
    
    