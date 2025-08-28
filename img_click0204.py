import pyautogui,os,sys,time,math,pyperclip,subprocess,unicodedata,csv # type: ignore
import tkinter as tk
from tkinter import messagebox
import threading
from tkinter import ttk

# 全角の文字列
FULLWIDTH_DIGITS = "０１２３４５６７８９"
FULLWIDTH_ALPHABET = "ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
FULLWIDTH_PUNCTUATION = "！＂＃＄％＆＇（）＊＋，－．／：；＜＝＞？＠［＼］＾＿｀｛｜｝～　"
FULLWIDTH_ALPHANUMERIC = FULLWIDTH_DIGITS + FULLWIDTH_ALPHABET  # 英数字
FULLWIDTH_ALL = FULLWIDTH_ALPHANUMERIC + FULLWIDTH_PUNCTUATION  # 英数字、記号

# 半角の文字列
HALFWIDTH_DIGITS = "0123456789"
HALFWIDTH_ALPHABET = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
HALFWIDTH_PUNCTUATION = "!\"#$%&'()*+,-./:;<=>?@[\]^_`{|}~ "
HALFWIDTH_ALPHANUMERIC = HALFWIDTH_DIGITS + HALFWIDTH_ALPHABET  # 英数字
HALFWIDTH_ALL = HALFWIDTH_ALPHANUMERIC + HALFWIDTH_PUNCTUATION  # 英数字、記号

bg_color_all = "#f6e5cc" #エクルベージュ ecru beige
bg_color = "#946c45" #"#9dc04c" #全体のボタン色:leaf green
bg_color2 = "#393f4c" #藍鉄 あいてつ
bg_color3 = "#7fffd4" #aquamarine
bg_color4 = "#ffa07a" #lightsalmon
bg_color5 = "#afeeee" #paleturquoise
hl_color= "#20b2fa" #https://www.colordic.org/
abg_color = "#fbdac8"   #シェルピンク shell pink ラジオボタン選択時背景色
fg_color = "#ffffff"
font_family = "Ubuntu"
cursor = "coffee_mug"
stop_posion = False
input_num = "" #仮の受付番号
repair_posion = 'start'

button_width = 18
csv_encoding = "utf-8"

oda = "image/oda.png"                   #各ボタンをイメージ化
enter = "image/enter.png"
#nfc = "image/nfc.png"
rp = "image/rp.png"
rp_check = "image/rpcheck.png"
rp_ok = "image/rp_ok.png"
stop = "image/stop.png"
go = "image/go.png"
print_rp = "image/print_rp.png"
cts_stop = "image/cts_stop.png"
nfc = "image/nfc2.png"
rsi_cancel = "image/rsi_cancel.png"
test_ok = "image/test_ok.png"
exchg = "image/exchg.png"
analysis = "image/analysis.png"
repair_type = "image/repair_type.png"
ccn = "image/ccn_tag.png"
repair_tag = "image/repair_tag.png"
break_tag = "image/break_tag.png"
repair_start ="repair_start"
parts_start = "image/parts_start.png"
parts_end = "image/parts_end.png"
scan_start = "image/scan_start.png"
repair_completed_img = "image/repair_completed.png"
check_oda_ok = "image/check_oda_ok.png"
oda_finish = "image/oda_finish.png"
back_to_tsp = "image/back_to_tsp.png"

repair_time = 0
break_time = 0

model_name = ""

tab_split_list = []
missing_parts_list = []
embedded_parts_list = []
repair_standard_time_list = []

analysis_code =["7001","7005"]

btn_comment = [
    "ODAをスタートさせる",
    "分析完了『"+analysis_code[0]+"』",
    "ﾁｬｯｸ詰『"+analysis_code[0]+"』",
    "ﾊﾞｯﾃﾘ交換『"+analysis_code[0]+"』",
    "充電器交換『"+analysis_code[0]+"』",
    "貸出異常無『"+analysis_code[1]+"』",
    "本体異常無『"+analysis_code[1]+"』"
    ]

sleep_time = 0.3
wait_time = 2.0
btn_padding = 6

#同梱しない場合はコードで以下のpathにすれば、exeファイルと同じフォルダを指定できます。 
script_dir = os.path.dirname(os.path.abspath(sys.argv[0])) 
file_path = os.path.join(script_dir, "data/repair_code.csv")
repairtime_data_path = os.path.join(script_dir, "data/repair_time.csv")
set_btn_path = os.path.join(script_dir, "image/setting_btn_img.png")

#set_btn_path = os.path.join(script_dir,oda)

btn_info = []
My_Code = []

def csv_data_read():
    csv_encoding = "utf-8"
    with open(file_path, encoding=csv_encoding) as f:
        reader = csv.reader(f)
        for row in reader:
            btn_info.append(row)

def csv_repairtime_data_read():
    csv_encoding = "utf-8"
    with open(repairtime_data_path, encoding=csv_encoding) as f:
        reader = csv.reader(f)
        for row in reader:
            repair_standard_time_list.append(row)

def start_check():
    time.sleep(wait_time)
    drag_position =check(analysis)
    #drag_position =check("repair_start")
    global model_name
    model_name = drag("model_name",drag_position)
    global thread_get_time
    thread_get_time = threading.Thread(target=standard_time_get, args=(model_name,))
    thread_get_time.start()
    base.title("ボタンをクリック")

def input_start(): #ボタンを押したときに前のウィンドウに戻る
    time.sleep(sleep_time)
    pyautogui.hotkey('alt', 'escape')

def start_oda(): # "ODAをスタートさせる"
    input_start()
    time.sleep(sleep_time)
    click(go)
    time.sleep(wait_time)
    click(oda)

def ok_oda_nuron_battery():
    time.sleep(wait_time*3)
    check(check_oda_ok)
    click(oda_finish)
    time.sleep(wait_time)
    click(back_to_tsp)


def nfc_click():
    input_start()
    time.sleep(sleep_time)
    click(nfc)  

def stop_window_display(click):
    if click == "open":
        stop_window.deiconify()
    else:
        stop_window.withdraw()

def stop_click(img,check_img):
    print(stop)
    global top_btn
    stop_window.withdraw()
    base.after(200, click, img)
    base.after(400, stop_window_display,"open")
    base.after(500, check,check_img)
    stop_btn.config(text="Wait...")
    
def click(img):
    set_btn_path = os.path.join(script_dir, img)
    print(img)
    img_look = False
    loop=0
    loop2 =0
    while img_look == False:
        try:
            p = pyautogui.locateOnScreen(set_btn_path, confidence=.5)
            print(p)
            global x,y
            x, y = pyautogui.center(p)
            print("x,y=",x, y)
            print("x=",x)
            print("y=",y)
            if img == test_ok:
                pyautogui.click(x+120, y)
                time.sleep(sleep_time)
                return
            pyautogui.click(x, y)
            time.sleep(sleep_time)
            break
        except pyautogui.ImageNotFoundException:
            if loop > 3:
                print("見つからず")
                break
            print(loop,"回目まだ見つかりません",img)
            time.sleep(1)
            loop += 1
            continue
def drag(name,position):
    if name == "model_name":
        print(position)
        pyautogui.moveTo(position)
        pyautogui.drag(150, 0, 1, button='left')
        #pyautogui.hotkey('command','c') #mac
        pyautogui.hotkey('ctrl','c') #windows
        model_name =pyperclip.paste()
        print("コピーデータ",model_name)
        model_name = 'SID 4-A22 01'
        print(model_name)
        return model_name

def analysis_completed(btn,comment,type,text_code):
    thread3 = threading.Thread(target=analysis_completed_thread, args=(comment,type,btn,text_code,))
    thread3.start()
    btn.config(text="実行中")
    btn_cantclick()

def analysis_completed_thread(comment,type,btn,text_code):
    input_start()
    time.sleep(sleep_time)
    pyautogui.typewrite(comment)
    pyautogui.press('enter')
    time.sleep(wait_time)
    click(go)
    time.sleep(wait_time)
    pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.typewrite(type)
    time.sleep(wait_time)
    click(go)
    btn.config(text=text_code)
    btn_okclick()
    repair_code_btn_start()

def btn_display(thread):
    
    thread.join()
    btn_okclick()

def nfc_and_rp(btn): #NFCの書き込みボタンを押すとレイティングプレートも印刷される
    btn.config(text="実行中")
    thread3 = threading.Thread(target=nfc_and_rp_thread)
    thread3.start()
    btn_cantclick()
    
def nfc_and_rp_thread():
    click(nfc)
    time.sleep(wait_time)
    check(print_rp)
    time.sleep(wait_time)
    check(rp_check)
    check(rp_ok)
    nfc_and_rp_btn.config(text="NFCの後RP")
    btn_okclick()

def change_wait_time():
    global wait_time
    wait_time = float(wait_time_setting.get())
    print("wait_timeが",wait_time,"秒に変更されました。")

def check(img): #画像が出るかを確認
    if img == "repair_start":
        set_btn_path2 = os.path.join(script_dir, analysis)
        set_btn_path = os.path.join(script_dir, repair_tag)
    else:
        set_btn_path = os.path.join(script_dir, img)
    print("認識中",img)
    img_look = False
    loop=0
    loop2 =0
    while img_look == False:
        try:
            if pyautogui.locateOnScreen(set_btn_path):
                p = pyautogui.locateOnScreen(set_btn_path, confidence=.5)
                print(p)
                x, y = pyautogui.center(p)
                print("認識しました",img)
                global stop_posion
                global break_time
                if img == analysis or img == repair_tag:
                    if stop_posion == True:
                        print("stop_posion == True",img)
                        stop_btn.config(text="Stop!!")
                        stop_posion = False
                        break_time = 0
                        print(break_time)
                    else:
                        position = [x+50, y]
                        return position
                        check(ccn) #まだ実装しない
                elif img == "repair_start":
                    print("修理開始します") 
                    if stop_posion == True:
                        stop_posion = False
                        break_time = 0
                        stop_btn.config(text= "Stop!!",command=lambda:stop_click(stop,break_tag))
                        print(break_time)
                elif img == break_tag:
                    if stop_posion == False:
                        stop_posion = True
                        print("ブレーク中です")
                        stop_btn.config(text= "START",command=lambda:stop_click(stop,repair_start))
                        break
                elif img == repair_type:
                    #ccn_window=tk.Toplevel(base)
                    tk.messagebox.showinfo(title="修理情報", message="修理です")
                elif img == ccn:
                    tk.messagebox.showinfo(title="ccn", message="CCNです")
                else:
                   pyautogui.click(x, y) 
                img_look = True
                print("x,y=",x, y)
                print("x=",x)
                print("y=",y)
        # 「except pg.ImageNotFoundException」で画像認識できなかった時の動作を指定
        except pyautogui.ImageNotFoundException:
            if loop > 3:
                if img == "repair_start" and loop2 == 0:
                    set_btn_path = set_btn_path2
                    print("修理タグが見つからず、分析タグを探します")
                    loop = 0
                    loop2 += 1
                    continue
                if img == "repair_start" and loop2 == 1:
                    print("分析タグが見つからず、ブレイクタイム継続中")
                    stop_btn.config(text="START")
                    break
                stop_btn.config(text="STOP!!")
                break
            print(loop,"回目まだ見つかりません",img)
            time.sleep(3)
            loop += 1
            continue

def btn_cantclick():
    for btn_create in btn_list:
        btn_create.config(state="disabled")

def btn_okclick():
    for btn_create in btn_list:
        btn_create.config(state="normal")

def my_find(l, x):
    if x in l:
        return l.index(x)
    else:
        return -1

def standard_time_get(model_name):
    csv_repairtime_data_read()
    print(model_name)
    print("timedata",repair_standard_time_list)
    i=0
    for model_name_time in repair_standard_time_list:
        get_data = my_find(model_name_time,model_name)
        print("get_data",get_data)
        if  get_data != -1:
            min_data = int(model_name_time[1])
            break
    global sed_data
    sed_data = min_data*60
    print(sed_data)
    return sed_data

#repair_standard_time = sed_data #ココのデータをどこからか持ってくる

def repair_timer():
    print("タイマースタート")

    global break_time
    global repair_posion
    global repair_time
    thread_get_time.join()
    progress.config(maximum=sed_data)
    #thread = threading.Thread(start_progress())
    while True:
        if stop_posion == False and repair_posion == 'completed':
            print("修理完了のためタイマー終了")
            break
        elif stop_posion == False:
            time.sleep(0.99)
            progress.start(sed_data)
            global repair_time
            progress['value'] = repair_time
            #base.update_idletasks()
            if stop_posion == True:
                break_time = 0
                continue
            min = math.floor(repair_time/60)
            second = repair_time-(60*min)
            if second < 10:
                second = str(0)+str(second)
            total_time = min,":",second
            repair_time += 1
            # progress['value'] = repair_time
            base.after(10,repair_time_label.config(text=total_time,fg=bg_color))
            #repair_time_label.config(text=total_time,fg='black')
            if stop_posion == True:
                break_time = 0
                continue
        else:
            time.sleep(5)
            repair_time_label.config(text="BreakTime")
            if stop_posion == False:
                print("stop_posion",stop_posion)
                continue
            #min = math.floor(break_time/60)
            #second = break_time-(60*min)
            #if second < 10:
                second = str(0)+str(second)
                print(second)
            #break_total_time = min,":",second
            base.after(5000,repair_time_label.config(text=total_time,fg='blue'))
            #break_time += 1
            if stop_posion == False:
                print("stop_posion",stop_posion)
                continue

def get_img_position(img_path): #imageの位置情報を取得
    get_img_path = os.path.join(script_dir, img_path)
    img_look = False
    loop = 0
    while img_look == False:
        try:
            if pyautogui.locateOnScreen(get_img_path):
                p = pyautogui.locateOnScreen(get_img_path, confidence=.5)
                print(p)
                x, y = pyautogui.center(p)
                print("認識しました",img_path)
                img_look = True
                print("x,y=",x, y)
                print("x=",x)
                print("y=",y)
                return p
        # 「except pg.ImageNotFoundException」で画像認識できなかった時の動作を指定
        except pyautogui.ImageNotFoundException:
            if loop > 3:
                break
            print(loop,"回目まだ見つかりません",img_path)
            time.sleep(3)
            loop += 1
            continue

def get_missing_parts_data():
    img_start_position = get_img_position(parts_start) #戻り値は4変数のタプル（left, top, width, height）です
    img_end_position = get_img_position(parts_end)     #https://qiita.com/suipy/items/15fe275d649c19a51a3e
    print(img_start_position)
    print(img_end_position)
    start_x,start_y = img_start_position[0],img_start_position[1]
    print(start_x,start_y)
    end_x,end_y = img_end_position[0]+img_end_position[2],img_end_position[1]
    print(end_x,end_y)
    pyautogui.click(start_x,start_y)
    time.sleep(5)
    pyautogui.dragTo(end_x,end_y, 1, button='left')
    pyautogui.hotkey('alt', 'escape')
    #with pyautogui.hold('command'): #mac
    with pyautogui.hold('ctrl'): #windows
        pyautogui.press('c')
    clip_str_check()
    click_parts_mail(input_num)
    
def clip_str_check():
    clip_str = pyperclip.paste()
    print(clip_str)

    n_list = []
 
    rn_split=clip_str.split('\r\n')
  
    for row in rn_split:
        n_list.append(row)
    print(n_list)
    
    for one in n_list:
        tab_split=one.split('\t')
        print(tab_split)
        if tab_split[2] != '-' or len(tab_split) < 1:
            if int(tab_split[2]) >= 0:
                tab_split[2] = int(tab_split[2])
            if int(tab_split[4]) >= 0:    
                tab_split[4] = int(tab_split[4])
            if tab_split[2] != "-" and tab_split[2] == 0 and tab_split[4] != 0:
                missing_parts_list.append(tab_split)
            elif tab_split[2] != "-" and tab_split[2] > 0:
                embedded_parts_list.append(tab_split)
        tab_split_list.append(tab_split)
       
            

    print(tab_split_list)
    print("missing_parts_list",missing_parts_list)
    print("embedded_parts_list",embedded_parts_list)

def remove_duplicates(lst):
    return lst(set(lst))

def clip_time_check():
    clip_str = pyperclip.paste()
    print(clip_str)
    n_list = []
    time_list=[]
    rn_split=clip_str.split('\r\n')
  
    for row in rn_split:
        n_list.append(row)
    print(n_list)
    
    for one in n_list:
        tab_split=one.split('\t')
        time_list.append([tab_split[2],tab_split[20]])
        print(tab_split[2],tab_split[20])
    time_list(set(time_list))
    print(time_list)

def click_parts_mail(num):
    print(num)
    msg_html = os.path.join(script_dir+"/msg/msg.html")
    with open(msg_html,'r', encoding=csv_encoding) as file:
        html = file.read()
    file.closed

    missing_html_contents = ''
    embedded_html_contents = ''

    print(missing_parts_list)
    if len(missing_parts_list) > 1:
        for missing_part in missing_parts_list:
            print(missing_part)
            missing_html_contents += missing_part[0]
            missing_html_contents += " &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;"
            missing_html_contents += str(missing_part[4])
            missing_html_contents += "<br>"
    elif len(missing_parts_list) == 1:
        missing_html_contents += missing_parts_list[0][0]
        missing_html_contents += " &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;0 &nbsp; &nbsp; &nbsp;"
        missing_html_contents += str(missing_parts_list[0][4])
        missing_html_contents += "<br>"

    if len(embedded_parts_list) > 1:
        for embedded_part in embedded_parts_list:
            embedded_html_contents += embedded_part[0]
            embedded_html_contents += " &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;"
            embedded_html_contents += str(embedded_part[4])
            embedded_html_contents += "<br>"
    elif len(embedded_parts_list) == 1:
        embedded_html_contents += embedded_parts_list[0][0]
        embedded_html_contents += " &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;"
        embedded_html_contents += str(embedded_parts_list[0][4])
        embedded_html_contents += "<br>"

    missing_parts_list.clear() #空にしないと次の使うときに
    embedded_parts_list.clear()#今回のものが入った状態からさらに加えて行くことになる

    page_data = {}
    page_data['num'] = num
    page_data['missing'] = missing_html_contents
    page_data['embedded'] = embedded_html_contents

    print(missing_html_contents)
    print(embedded_html_contents)

    print(html)
    for key, value in page_data.items():
        html = html.replace('{% ' + key + ' %}', value)
    print(html)
    #pyperclip.copy(html)
    print("コピーしました")
    msg_output_html = os.path.join(script_dir+"/msg/msg_output.html")
    with open(msg_output_html, 'w',encoding=csv_encoding) as f_out:
        f_out.write(html)
    pyperclip.copy(html)

    #subprocess.Popen(msg_output_html)
    #subprocess.run(msg_output_html)

def cursor_open():
    print("open")
    #cts_btn.config(cursor="spraycan",bg=bg_color),adm_btn.config(cursor="arrow",bg=bg_color),ccn_btn.config(cursor="spider",bg=bg_color),adm_btn.config(cursor="trek",bg=bg_color),parts_btn.config(cursor="man",bg=bg_color),sap_btn.config(cursor="mouse",bg=bg_color)

def num_input(event):
    global input_num
    global repair_posion
    input_num_window.withdraw()
    thread2 = threading.Thread(target=repair_timer)
    input_num = num_entry.get()
    repair_posion = 'start'
    for one_word in input_num:
        if unicodedata.east_asian_width(one_word) == 'Na' and len(input_num) == 8:
            if input_num.isdecimal() :
                text1 = "No."
                text2 = " Inputed"
                base.title(text1+input_num+text2)
                num_label.config(text=("Number.",input_num))
                num_entry.config(bg="white")
                num_entry.delete(0, tk.END)
                cursor_open()
                click(scan_start)
                pyautogui.write(input_num)
                pyautogui.hotkey('enter')
                start_check()
                thread2.start()
                break
        elif unicodedata.east_asian_width(one_word) == 'F' and len(input_num) == 8:
            if input_num.isdecimal() :
                print("全角です")
                conv_map = str.maketrans(FULLWIDTH_ALPHANUMERIC, HALFWIDTH_ALPHANUMERIC)
                input_num=input_num.translate(conv_map)
                text1 = "No."
                text2 = " Inputed"
                base.title(text1+input_num+text2)
                num_label.config(text=("Number.",input_num))
                num_entry.config(bg="white")
                num_entry.delete(0,tk.END)
                cursor_open()
                click(scan_start)
                pyautogui.write(input_num)
                pyautogui.hotkey('enter')
                thread2.start()
                break
        else :
            print("数字以外です")
            num_entry.delete(0,tk.END)
            num_entry.insert(tk.END,"受付番号を")
            #pyautogui.hotkey('command', 'a') #mac
            pyautogui.hotkey('ctrl', 'a') #windows
            break

repair_completed_time_list = []

def repair_completed():
    repair_completed_btn_window.withdraw()
    repair_code_btn_window.withdraw()
    time.sleep(0.2)
    click(repair_completed_img)
    global repair_posion
    global repair_time
    repair_posion = 'completed'
    min = math.floor(repair_time/60)
    second = repair_time-(60*min)
    if second < 10:
        second = str(0)+str(second)
    total_time = str(min)+":"+str(second)
    repair_completed_time_list.append([input_num,model_name,total_time])
    print(repair_completed_time_list)
    repair_time = 0
    base.deiconify()
    input_num_window.deiconify()
    pyautogui.moveTo(600, 150)
    num_entry.focus_set()

def repair_tiamer_go():
    thread2 = threading.Thread(target=repair_timer)

def repair_code_btn_list():
    print("チェック")

def repair_testok():
    pyautogui.hotkey('alt', 'escape')
    click(go)
    time.sleep(wait_time)
    pyautogui.click(x=780, y=540)
    i = 0
    while i < 2:
        time.sleep(sleep_time)
        pyautogui.hotkey('tab')
        i += 1
    pyautogui.press('space')
    time.sleep(wait_time)

"""def repair_testok():
    posi = [780,540]
    pyautogui.hotkey('alt', 'escape')
    click(go)
    base.after(2000,pyautogui.click,posi)
    base.after(2300,pyautogui.hotkey,'tab')
    base.after(2600,pyautogui.hotkey,'tab')
    base.after(4600,pyautogui.press,'space')
    #time.sleep(wait_time)
    #pyautogui.click(x=780, y=540)
    #i = 0
    #while i < 2:
        #time.sleep(sleep_time)
        #pyautogui.hotkey('tab')
        #i += 1
    #pyautogui.press('space')
    #time.sleep(wait_time)"""

def input_text0(abc):
    repair_testok()
    sleep_time = 0.3
    time.sleep(sleep_time)
    time.sleep(sleep_time)
    pyautogui.typewrite(abc) 
    time.sleep(sleep_time)
    pyautogui.press('enter')
    repair_completed_btn_window.deiconify()

def input_text(abc,code):
    repair_testok()
    sleep_time = 0.3
    time.sleep(sleep_time)
    time.sleep(sleep_time)
    pyautogui.typewrite(code) 
    time.sleep(sleep_time)
    i = 0
    while i < 1:
        pyautogui.hotkey('shift', 'tab')
        i += 1
    pyautogui.typewrite(abc) 
    time.sleep(sleep_time)
    pyautogui.press('enter')
    j=0
    global repair_code_list
    for repair_config in repair_code_list:
        repair_config.config(bg=bg_color,activebackground=abg_color, command=lambda repair_config=repair_config,j=j:input_text0(btn_info[j][1]))
        if j < len(btn_info)-1:
                j += 1
    l2["text"] = ""
    accessory_btn.config(bg="darkkhaki",command=lambda:input_text0("1026"))
    accessory_exchange_btn.config(bg="salmon",command=lambda:input_text0("1213"))
    repair_completed_btn_window.deiconify()

def input_text2(abc):
    repair_testok()
    sleep_time = 0.3
    time.sleep(sleep_time)
    time.sleep(sleep_time)
    pyautogui.typewrite(abc) 
    pyautogui.press('enter')
    j=0
    global repair_code_list
    for repair_config in repair_code_list:
        repair_config.config(bg=bg_color,activebackground=abg_color, command=lambda repair_config=repair_config,j=j:input_text0(btn_info[j][1]))
        if j < len(btn_info)-1:
                j += 1
    l2["text"] = ""
    rent_btn.config(state="disabled")
    repair_completed_btn_window.deiconify()

def input_text_all(abc):
    sleep_time = 0.3
    time.sleep(sleep_time)
    pyautogui.hotkey('alt', 'escape')
    pyautogui.typewrite(abc)
    i = 0
    while i < 12:
        pyautogui.hotkey('tab')
        i += 1
    time.sleep(2)
    j = 0
    while i < 2:
        pyautogui.hotkey('shift', 'tab')
        j += 1
    pyautogui.typewrite(7001) 
    time.sleep(1)
    k = 0
    while k < 5:
        pyautogui.hotkey('tab')
        k += 1
    pyautogui.press('space')

def setting_window(list):
    sw_list = list
    setting_window = tk.Toplevel(base)
    setting_window.title('Botton Setting')
    setting_info = tk.Label(setting_window, text ="ボタンの表示内容と添付内容を入力してください",bg=bg_color,font=("Ubutu", 12))
    setting_info.pack(pady=12, padx=12)
    btn_setting_list = []
    frame_list = []
    for btn_change in sw_list:
        frame_top = tk.Frame(setting_window, pady=5, padx=5, relief=tk.RAISED, bd=2)
        frame_list.append(frame_top)
        entry = tk.Entry(frame_top,font=(font_family , "14", "bold"), bd=0,width=8)
        entry.insert(tk.END,btn_change[0])
        entry.pack(side=tk.LEFT,pady=5)
        entry2 = tk.Entry(frame_top,font=(font_family , "14", "bold"),bd=0,width=30)
        btn_setting_list.append([entry,entry2])
        entry2.insert(tk.END,btn_change[1])
        entry2.pack(pady=5)
        frame_top.pack()
    print(btn_setting_list)
    cancel = tk.Button(setting_window,font=(font_family , "12", "bold"), text="Cancel",bg=bg_color,activebackground=abg_color, command=setting_window.destroy)
    cancel.pack(side=tk.LEFT,padx=5,pady=5)
    check = tk.Button(setting_window,font=(font_family , "12", "bold"), text="Check",bg=bg_color,activebackground=abg_color, command=lambda:check_setting(setting_window,frame_list,setting_info,btn_setting_list,cancel,check))
    check.pack(side=tk.RIGHT,padx=5,pady=5)

def check_setting(window,frame,info,list,l_btn,r_btn):
    print(list)
    check_list = []
    for btn_generate in list:
        check_list.append([btn_generate[0].get(),btn_generate[1].get()])
    print(len(list))
    info.config(text = "下記の内容でよろしいでしょうか？")
    for btn in list:
        btn[0].destroy()
        btn[1].destroy()
    print(check_list)
    print(frame)
    i = 0
    for btn_generate in list:
        btn_gnt = tk.Button(frame[i],font=(font_family , "12"), text=check_list[i][0],bg=bg_color,activebackground=abg_color)
        btn_details = tk.Label(frame[i], text=check_list[i][1], bg="lightblue")
        btn_gnt.pack(side=tk.LEFT,pady=5)
        btn_details.pack(pady=5)
        if i < len(check_list):
            i += 1
    l_btn.config(text="Back",bg=bg_color,activebackground=abg_color, command=lambda:resetting_window(window,check_list))
    r_btn.config(text="Choose",bg=bg_color,activebackground=abg_color, command=lambda:choose_btn_data(window,check_list))

def resetting_window(window,list):
    window.destroy()
    setting_window(list)

def choose_btn_data(window,list):
    with open(file_path, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(list)  # 2次元配列を一度に書き込む
    
    update_btn_data(list)
    window.destroy()

def update_btn_data(list):
    global btn_info
    btn_info = list
    global repair_code_list 
    i = 0
    for update_btn in repair_code_list: #ボタンの更新（OnTrackNo.入力しないでの修理コードを入力）
        update_btn.config(text=btn_info[i][0],command=lambda update_btn=update_btn,i=i:input_text0(btn_info[i][1])) #ラムダ式にFor文を使う場合、repair_code=repair_codeのように入力が必要https://teratail.com/questions/374578
        print(update_btn)
        if i < len(btn_info)-1:
            i += 1
    print(repair_code_list)

def input(event):
    input_num = e2.get()
    print("e2.get",input_num)
    if len(input_num) == 9 and (input_num[0:1] == "4" or input_num[0:1] == "3"):
        l2["text"] = "On" + input_num        #set label text
        e2.delete(0, tk.END)
        i=0
        print("btn_infoのレングス",len(btn_info))
        print(btn_info[i])
        for repair_on_config in repair_code_list:
            repair_on_config.config(bg=bg_color2,activebackground=abg_color, command=lambda repair_on_config=repair_on_config,i=i:input_text(btn_info[i][1],input_num ))
            if i < len(btn_info)-1:
                i += 1
        
def change_padding():
    global btn_padding
    print(btn_padding_setting.get())
    print(btn_padding)
    btn_padding = int(btn_padding_setting.get())
    for pack_btn in repair_code_list:
        pack_btn.config(pady=btn_padding)

def repair_code_btn_start():
    repair_code_btn_window.deiconify()
    base.withdraw()
    l2.config(text="Ontrack入力欄")
    e2.focus_set()

def input_ontrack(event):
    input_num = e2.get()
    print("e1.get",input_num)
    if len(input_num) == 9 and (input_num[0:1] == "4" or input_num[0:1] == "3"):
        l2["text"] = "On" + input_num        #set label text
        e2.delete(0, tk.END)
        i=0
        print("btn_infoのレングス",len(btn_info))
        print(btn_info[i])
        for repair_on_config in repair_code_list:
            repair_on_config.config(bg=bg_color2,activebackground=abg_color, command=lambda repair_on_config=repair_on_config,i=i:input_text(btn_info[i][1],input_num ))
            if i < len(btn_info)-1:
                i += 1
        accessory_btn.config(bg="darkkhaki",command=lambda:input_text("1026",input_num ))
        accessory_exchange_btn.config(bg="salmon",command=lambda:input_text("1213",input_num ))
        rent_btn.config(command=lambda :input_text2(input_num),state='normal')
        
base = tk.Tk() #問題画面を作成
base.geometry("200x650")
base.geometry("+1720+350")
base.title("読み込み中")
base.attributes("-topmost", True) #最前面に配置
base.tk_setPalette(background=bg_color_all,	foreground=fg_color)
thread1 = threading.Thread(target=start_check)
#thread_get_time = threading.Thread(target=standard_time_get, args=(model_name,))
"red"
#thread1.start()


oda_btn = tk.Button(base,font=(font_family, "12"),text=btn_comment[0],bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=start_oda,width=button_width,cursor=cursor,relief=tk.RAISED)
oda2_btn = tk.Button(base,font=(font_family, "12"),text="ODAを完了まで",bg=bg_color,activebackground=abg_color,command=start_oda,width=button_width,cursor=cursor,relief=tk.RAISED)
cts_btn = tk.Button(base,font=(font_family, "12"),text="NFCクリック",bg=bg_color,activebackground=abg_color,command=nfc_click,width=button_width,cursor=cursor)
analysis_completed_btn = tk.Button(base,font=(font_family, "12"),text=btn_comment[1],bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:analysis_completed(analysis_completed_btn,"19920;372016;",analysis_code[0],btn_comment[1]),width=button_width,cursor=cursor)
analysis_chack_btn = tk.Button(base,font=(font_family, "12"),text=btn_comment[2],bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:analysis_completed(analysis_chack_btn,"19920;372016;OP2108;DC1267;OP299;DC9057;",analysis_code[0],btn_comment[2]),width=button_width,cursor=cursor)
analysis_battey_btn = tk.Button(base,font=(font_family, "12"),text=btn_comment[3],bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:analysis_completed(analysis_battey_btn,"372016;OP701;DC9037; ",analysis_code[0],btn_comment[3]),width=button_width,cursor=cursor)
analysis_cahrger_btn = tk.Button(base,font=(font_family, "12"),text=btn_comment[4],bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:analysis_completed(analysis_cahrger_btn,"372016;OP702;DC9037; ",analysis_code[0],btn_comment[4]),width=button_width,cursor=cursor)
analysis_noproblem_btn = tk.Button(base,font=(font_family, "12"),text=btn_comment[5],bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:analysis_completed(analysis_noproblem_btn,"19920;372016;OP420;DC9025; ",analysis_code[1],btn_comment[5]),width=button_width,cursor=cursor)
analysis_norepair_btn = tk.Button(base,font=(font_family, "12"),text=btn_comment[6],bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:analysis_completed(analysis_norepair_btn,"19920;372016;OP2108;DC1267; ",analysis_code[1],btn_comment[6]),width=button_width,cursor=cursor)
nfc_and_rp_btn = tk.Button(base,font=(font_family, "12"),text="NFCの後RP",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:nfc_and_rp(nfc_and_rp_btn ),width=button_width,cursor=cursor)
go_btn = tk.Button(base,font=(font_family, "12"),text="GO Click",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:click(go),width=button_width,cursor=cursor)
nfc_btn = tk.Button(base,font=(font_family, "12"),text="NFC Write",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:click(nfc),width=button_width,cursor=cursor)

#リペアタイマー
stop_window=tk.Toplevel(base)
stop_window.geometry("120x100+1078+394")
stop_window.attributes("-topmost", True) #最前面に配置
stop_btn = tk.Button(stop_window,font=(font_family, "12"),text="STOP!!",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:stop_click(stop,break_tag),width=10,cursor=cursor)
repair_time_label = tk.Label(stop_window,text="0:00",fg=bg_color)
# Progressbarの作成
progress = ttk.Progressbar(stop_window, orient="horizontal", length=120, mode="determinate",maximum=900)
print_btn = tk.Button(base,font=(font_family, "12"),text="print",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:click(print_rp),width=button_width,cursor=cursor)
ccn_btn = tk.Button(base,font=(font_family, "12"),text="CCNチェック",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:check(repair_type),width=button_width,cursor=cursor)
missing_parts_btn = tk.Button(base,font=(font_family, "12"),text="部品待ち情報",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=get_missing_parts_data,width=button_width,cursor=cursor)
parts_btn = tk.Button(base,font=(font_family, "12"),text="部品待ちメール",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=lambda:click_parts_mail(input_num),width=button_width,cursor=cursor)
#Label
num_label = tk.Label(base, width=25, bg="#5f6527")

repair_code_btn_window = tk.Toplevel(base) #修理コードボタン画面を作成
x =repair_code_btn_window.winfo_screenwidth()- repair_code_btn_window.winfo_width() #モニターのサイズを採って来ている
print(repair_code_btn_window.winfo_screenwidth())
print(repair_code_btn_window.winfo_screenheight())
window_height = str(math.floor(repair_code_btn_window.winfo_screenheight()*0.9))
window_width = '100x'
widow_size = window_width+window_height
repair_code_btn_window.geometry(widow_size)
repair_code_btn_window.geometry("+%d+%d" % (x-120, 35))
repair_code_btn_window.attributes("-topmost", True) #最前面に配置
repair_code_btn_window.withdraw()
#repair_code_btn_window.deiconify()activeforeground=bg_color,
img = tk.PhotoImage(file=set_btn_path)
repair_code_list = []
csv_data_read()
for repair_code in btn_info: #デフォルトのボタン作成（OnTrackNo.入力しないでの修理コードを入力）
    code_btn = tk.Button(repair_code_btn_window,font=(font_family , "12"), text=repair_code[0],bg=bg_color,activebackground=abg_color, command=lambda repair_code=repair_code:input_text0(repair_code[1]),width=10) #ラムダ式にFor文を使う場合、repair_code=repair_codeのように入力が必要https://teratail.com/questions/374578
    repair_code_list.append(code_btn)
    print(repair_code[1])
print(repair_code_list)
setting_btn = tk.Label(repair_code_btn_window,image=img,anchor=tk.SE)
setting_btn.bind("<Button-1>", lambda e:setting_window(btn_info))
for pack_btn in repair_code_list:
    pack_btn.pack(padx=btn_padding,pady=btn_padding)
intVar = tk.IntVar(value=btn_padding)
btn_padding_setting = tk.Spinbox(repair_code_btn_window,font=(font_family, "15"),from_=1,to=20,increment=1,width=2,command=change_padding,textvariable=intVar,bg=bg_color,buttonbackground = "#5f6527")
l2 = tk.Label(repair_code_btn_window,font=(font_family, "11"),text="Ontrack入力欄", width=20, bg="#5f6527")
l2.pack(pady=10)
#Entry
e2 = tk.Entry(repair_code_btn_window, font=(font_family, "15"),bd=0,width=15,fg=bg_color)
e2.bind("<Return>", input_ontrack)      #bind Retun key
e2.pack(padx=3)
accessory_btn = tk.Button(repair_code_btn_window,font=(font_family , "12"), text="付属異常無",bg="olive",activebackground=abg_color,command=lambda:input_text0("1026"),width=10)
accessory_exchange_btn = tk.Button(repair_code_btn_window,font=(font_family , "12"), text="本体交換",bg="tomato",activebackground=abg_color,command=lambda:input_text0("1213"),width=10)
rent_btn = tk.Button(repair_code_btn_window,font=(font_family , "12"), text="貸出機",bg="yellow",activebackground=abg_color,width=10)
accessory_btn.pack(padx=btn_padding,pady=btn_padding),accessory_exchange_btn.pack(padx=btn_padding,pady=btn_padding),rent_btn.pack(padx=btn_padding,pady=btn_padding)
setting_btn.pack(side=tk.BOTTOM,anchor=tk.SE,padx=5,pady=5) #設定ボタンを設置
btn_padding_setting.pack(side=tk.BOTTOM,anchor=tk.SW)
rent_btn.config(state="disabled")
e2.focus_set()

#受付番号入力Entry
input_num_window=tk.Toplevel(base)
input_num_window.title("受付番号入力してください")
input_num_window.geometry("400x60+950+100")
input_num_window.attributes("-topmost", True) #最前面に配置
#Entry
num_entry  = tk.Entry(input_num_window, bd=5,width=button_width*2,insertbackground=bg_color,fg=bg_color)
num_entry.bind("<Return>", num_input)      #bind Retun key

#修理完了ボタン
repair_completed_btn_window=tk.Toplevel(base)
repair_completed_btn_window.title("受付番号入力してください")
repair_completed_btn_window.geometry("80x60+900+700")
repair_completed_btn_window.attributes("-topmost", True) #最前面に配置
#repair_completed_btn_window.deiconify() #修理コードを入れた後に表示させる
repair_completed_btn_window.withdraw()


doubleVar = tk.DoubleVar(value=wait_time)
wait_time_setting = tk.Spinbox(base,font=(font_family, "15"),from_=0.0,to=20.0,increment=0.1,width=4,command=change_wait_time,textvariable=doubleVar,bg=bg_color,buttonbackground = "#5f6527")
#btn_list = [oda_btn,cts_btn,analysis_completed_btn,nfc_and_rp_btn,go_btn,nfc_btn,stop_btn,repair_time_label,print_btn,ccn_btn,missing_parts_btn,parts_btn]
btn_list = [oda_btn,analysis_completed_btn,analysis_chack_btn,analysis_battey_btn,analysis_cahrger_btn,analysis_noproblem_btn,analysis_norepair_btn,print_btn,missing_parts_btn,parts_btn]
for btn_create in btn_list:
    btn_create.pack(padx=6,pady=6)

timer_display_parts = [stop_btn,repair_time_label,progress]
for part_create in timer_display_parts:
    part_create.pack(pady=1)

missing_parts_btn = tk.Button(base,font=(font_family, "12"),text="クリップボードチェック",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=get_missing_parts_data,width=button_width,cursor="circle").pack(pady=6)
missing_parts_btn = tk.Button(base,font=(font_family, "12"),text="文章スプリット",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=clip_str_check,width=button_width,cursor="circle").pack(pady=6)
missing_parts_btn = tk.Button(base,font=(font_family, "12"),text="修理標準時間",bg=bg_color,activebackground=abg_color,activeforeground=bg_color,command=clip_time_check,width=button_width,cursor="circle").pack(pady=6)
num_label.pack(padx=6,pady=6)
num_entry.pack(padx=6,pady=6)
repair_completed_btn = tk.Button(repair_completed_btn_window,font=(font_family, "12"),text="修理完了",bg="Red",activebackground=abg_color,command=repair_completed,width=button_width,cursor="circle").pack(padx=6,pady=6)
wait_time_setting.pack(side=tk.BOTTOM,anchor=tk.SW)

num_entry.focus_set()

base.mainloop()
