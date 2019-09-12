from selenium import webdriver
from tkinter import Tk, Label, Entry, Button, StringVar
from tkinter.font import Font
import re
import time
import os
import sys
import pyperclip


# 文献番号（dnum）の入力を受け付ける
def interface():
    root = Tk()
    root.title(u'文献番号入力')
    root.geometry('950x250')
    root.configure(background='SlateGray1')
    font1 = Font(family=u'游明朝', size=12)
    label1 = Label(text=u'文献番号を入力してください。', font=font1, background='SlateGray1')
    label2 = Label(text=u'注意事項', font=font1, background='SlateGray1')
    label3 = Label(text=u'＊入力した文字をJ-PlatPatの簡易検索ボックスにそのまま代入して検索します。',
                   font=font1, background='SlateGray1')
    label4 = Label(text=u'＊文献を一意に特定できるように入力してください。\n'
                   '（例えば、単に2019-100xxxと入力した場合、'
                   '特願2019-100xxxと特開2019-100xxxが検索結果として返される場合があります。）',
                   font=font1, background='SlateGray1', justify='left')
    dnum = StringVar()
    input1 = Entry(width=30, font=font1, textvariable=dnum)

    # 送信ボタン押下時の処理
    def btn_click():
        text = dnum.get()
        if(text == ''):
            root2 = Tk()
            root2.title(u'alert')
            root2.geometry('300x300')
            root2.configure(background='SlateGray1')
            label5 = Label(root2, text=u'文献番号が入力されていません。',
                           font=font1, background='SlateGray1')
            label5.pack(expand=1)
            root2.mainloop()
        else:
            root.destroy()

    button1 = Button(text=u'送信', font=font1, command=btn_click)
    label1.place(x=5, y=5)
    input1.place(x=5, y=40)
    button1.place(x=300, y=33)
    label2.place(x=5, y=75)
    label3.place(x=5, y=110)
    label4.place(x=5, y=145)

    root.mainloop()
    text = dnum.get()
    return text


# 検索結果を一つに特定できない場合の処理
def alert(result):
    root3 = Tk()
    root3.title(u'alert')
    root3.geometry('300x300')
    root3.configure(background='SlateGray1')
    font1 = Font(family=u'游明朝', size=12)
    label6 = Label(root3, text=u'１件もヒットしませんでした。',
                   font=font1, background='SlateGray1')
    label7 = Label(root3, text=u'検索結果が取得できませんでした。\n'
                   '検索ボタン押下後十分に時間をとってください。',
                   font=font1, background='SlateGray1')
    label8 = Label(root3, text=u'拒絶理由通知が未登録です。',
                   font=font1, background='SlateGray1')
    label9 = Label(root3, text=u'複数の文献がヒットしました。',
                   font=font1, background='SlateGray1')
    label10 = Label(root3, text=u'リンクを張れる引用文献がありません。',
                    font=font1, background='SlateGray1')
    label11 = Label(root3, text=u'処理が完了しました。',
                    font=font1, background='SlateGray1')
    if('(0)' in result):
        label6.pack(expand=1)
    elif('(-)' in result):
        label7.pack(expand=1)
    elif(result == 'NoOfficeAction'):
        label8.pack(expand=1)
    elif(result == 'NoCitation'):
        label10.pack(expand=1)
    elif(result == 'Complete'):
        label11.pack(expand=1)
    else:
        label9.pack(expand=1)
    root3.mainloop()
    driver.quit()


# 入力された文献番号（変数dnum）のヒット件数を取得する
def hitnumber(text):
    """
    入力BOX : 's01_srchCondtn_txtSimpleSearch'
    検索ボタン : 's01_srchBtn_btnSearch'
    ヒット件数 : 'mat-tab-label-0-0'
    """

    driver.get('https://www.j-platpat.inpit.go.jp/')
    driver.find_element_by_id('s01_srchCondtn_txtSimpleSearch').send_keys(text)
    driver.find_element_by_id('s01_srchBtn_btnSearch').click()
    time.sleep(2)
    result = driver.find_element_by_id('mat-tab-label-0-0').text
    return result


# 検索結果のボタンを押して拒絶理由通知を取得する
def oagetter():
    """
    経過情報 : 'patentUtltyIntnlNumOnlyLst_tableView_progReferenceInfo0'
    """

    driver.find_element_by_id(
        'patentUtltyIntnlNumOnlyLst_tableView_progReferenceInfo0').click()
    driver.switch_to.window(driver.window_handles[-1])
    lst = driver.find_elements_by_xpath(
        u"(//a[contains(text(), '拒絶理由通知書')])")
    if(len(lst) == 0):
        return None
    lst[-1].click()
    driver.switch_to.window(driver.window_handles[-1])
    OAText = driver.find_element_by_tag_name('pre').text
    driver.close()
    driver.switch_to.window(driver.window_handles[-1])
    driver.close()
    return OAText


# 拒絶理由通知書のhtmlとフォルダを作成する
def filecreator(OAText, dnum1, dirpath):
    os.makedirs(dirpath, exist_ok=True)
    filepath = os.path.join(dirpath, '%sOA.html' % dnum1)
    with open(filepath, 'wt') as f:
        f.write('<!DOCTYPE>\n<html>\n<pre>\n' +
                OAText+'\n</pre>\n</html>')
    return filepath


# 引用文献に特許文献が存在したら取得する
def doc_catcher(OAText):
    match = re.search(r'適用条文.*?((\S)\2{8,}|先行技術文献調査結果)', OAText, re.DOTALL)
    if(match is None):
        alert('NoCitation')
        sys.exit()
    else:
        p_cited = re.findall(
            '[ 　０-９]{1,2}.([特実再][許公開願表]{1,2}[昭平令]?[0-9０-９－ー-]+)', match.group())
        print('日本の引用特許文献は', p_cited)
        usp_cited = re.findall(
            '[ 　０-９]{1,2}.米国特許第([0-9０-９]+)号', match.group())
        print('米国の引用特許文献は', usp_cited)
        return p_cited, usp_cited


# 日本特許文献に対してリンクを取得しhtmlにリンクを張る
def jp_doc_linker(p_cited, filepath):
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)
    driver.find_element_by_id(
        's01_srchCondtn_txtSimpleSearch').clear()
    for pdoc in p_cited:
        driver.find_element_by_id(
            's01_srchCondtn_txtSimpleSearch').send_keys(pdoc+' ')
    driver.find_element_by_id('s01_srchBtn_btnSearch').click()
    with open(filepath) as f:
        filedata = f.read()
    for pnum in range(len(p_cited)):
        time.sleep(1)
        driver.find_element_by_id(
            'patentUtltyIntnlNumOnlyLst_tableView_url'+str(pnum)).click()
        driver.switch_to.window(driver.window_handles[-1])
        driver.find_element_by_id('btnClose').click()
        driver.switch_to.window(driver.window_handles[-1])
        filedata = filedata.replace(p_cited[pnum], '<a href='+pyperclip.paste() +
                                    ', target=_blanck>'+p_cited[pnum]+'</a>')
    with open(filepath, 'w') as f:
        f.write(filedata)


# 米国特許文献に対してリンクを取得しhtmlにリンクを張る
def us_doc_linker(usp_cited, filepath):
    driver.switch_to.window(driver.window_handles[0])
    driver.get('https://www.j-platpat.inpit.go.jp/p0000')
    for i, uspdoc in enumerate(usp_cited):
        time.sleep(1)
        driver.find_element_by_id(
            "p00_srchCondtn_btnDocNoInputCountry"+str(i)).click()
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(1)
        driver.find_element_by_xpath(
            u"(//a[contains(text(), 'アメリカ(US)')])").click()
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(1)
        driver.find_element_by_id(
            "p00_srchCondtn_btnDocNoInputType"+str(i)).click()
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(1)
        driver.find_element_by_xpath(
            u"(//a[contains(text(), '特許番号(A/B)')])").click()
        driver.switch_to.window(driver.window_handles[-1])
        time.sleep(1)
        driver.find_element_by_id(
            "p00_srchCondtn_txtDocNoInputNo"+str(i)).send_keys(uspdoc)
    driver.find_element_by_id(
        "p00_searchBtn_btnDocInquiry").click()
    with open(filepath) as f:
        filedata = f.read()
    for uspnum in range(len(usp_cited)):
        time.sleep(1)
        driver.find_element_by_id(
            "patentUtltyIntnlSimpleBibLst_tableView_url"+str(uspnum)).click()
        driver.switch_to.window(driver.window_handles[-1])
        driver.find_element_by_id('btnClose').click()
        driver.switch_to.window(driver.window_handles[-1])
        filedata = filedata.replace('米国特許第'+usp_cited[uspnum], '<a href='+pyperclip.paste() +
                                    ', target=_blanck>米国特許第'+usp_cited[uspnum]+'</a>')
    with open(filepath, 'w') as f:
        f.write(filedata)


if __name__ == '__main__':

    # 検索結果が１件になるまでデータ入力を繰り返す。
    while(True):
        dnum1 = interface()
        if(dnum1):
            driver = webdriver.Chrome()
            result = hitnumber(dnum1)
            if('(1)' not in result):
                alert(result)
                continue
            else:
                break
        else:
            sys.exit()

    # フォルダを作成する際のpathを決定
    dirpath = os.path.join(os.getcwd(), '%sOA' % dnum1)

    # 拒絶理由通知書を取得する
    OAText = oagetter()

    if(OAText is None):
        alert('NoOfficeAction')
        sys.exit()
    else:
        # 拒絶理由通知書のhtmlファイルとフォルダを作成する
        filepath = filecreator(OAText, dnum1, dirpath)

        # 拒絶理由通知書から日本語特許文献のリストと外国語特許文献のリストを取得する
        p_cited, usp_cited = doc_catcher(OAText)

        if(p_cited):
            jp_doc_linker(p_cited, filepath)

        if(usp_cited):
            us_doc_linker(usp_cited, filepath)

        # 処理の完了を通知
        alert('Complete')
        sys.exit()
