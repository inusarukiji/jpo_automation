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


# １回目の拒絶理由通知が取得できたらhtmlファイルに保存する
def oagetter(Doc, dirpath):
    """
    経過情報 : 'patentUtltyIntnlNumOnlyLst_tableView_progReferenceInfo0'
    """

    driver.find_element_by_id(
        'patentUtltyIntnlNumOnlyLst_tableView_url0').click()
    driver.switch_to.window(driver.window_handles[-1])
    driver.find_element_by_id('btnClose').click()
    url = pyperclip.paste()
    driver.find_element_by_id(
        'patentUtltyIntnlNumOnlyLst_tableView_progReferenceInfo0').click()
    driver.switch_to.window(driver.window_handles[-1])
    lst = driver.find_elements_by_xpath(
        u"(//a[contains(text(), '拒絶理由通知書')])")
    if(len(lst) == 0):
        driver.close()
        return None, url
    lst[0].click()
    driver.switch_to.window(driver.window_handles[-1])
    OAText = driver.find_element_by_tag_name('pre').text
    driver.close()
    driver.switch_to.window(driver.window_handles[-1])
    driver.close()
    with open(os.path.join(dirpath, '%sOA.html' % Doc), 'wt') as f:
        f.write('<!DOCTYPE>\n<html>\n<pre>\n'+OAText+'\n</pre>\n</html>')
    return OAText, url


if __name__ == '__main__':

    # 検索結果が１件になるまでデータ入力を繰り返す。
    while(True):
        docname = interface()
        if(docname):
            driver = webdriver.Chrome()
            result = hitnumber(docname)
            if('(1)' not in result):
                alert(result)
                continue
            else:
                break
        else:
            sys.exit()

    index = 0
    ancestor = [[docname]]
    url_dictionary = {}
    dictionary = {}
    dirpath = os.path.join(os.getcwd(), '%sOA' % docname)
    os.makedirs(dirpath, exist_ok=True)
    # 引用文献の親をたどる
    while(ancestor[index]):
        ancestor.append([])
        for Doc in ancestor[index]:
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(1)
            driver.find_element_by_id(
                's01_srchCondtn_txtSimpleSearch').clear()
            driver.find_element_by_id(
                's01_srchCondtn_txtSimpleSearch').send_keys(Doc)
            driver.find_element_by_id('s01_srchBtn_btnSearch').click()
            time.sleep(2)
            # 拒絶理由通知を取得
            OAText, url = oagetter(Doc, dirpath)
            url_dictionary[Doc] = url
            if(OAText is None):
                dictionary[Doc] = []
                continue
            else:
                # 主引用文献を取得
                m = re.search(r'適用条文.*?((\S)\2{8,}|先行技術文献調査結果)',
                              OAText, re.DOTALL)
                if(m is None):
                    dictionary[Doc] = []
                    continue
                else:
                    preancestor = re.findall(
                        '[ 　１1]{1,2}.([特実再][許公開願表]{1,2}[昭平令]?[0-9０-９－ー-]+)',
                        m.group())
            print(preancestor)
            ancestor[index+1].extend(preancestor)
            dictionary[Doc] = preancestor
        # ancestor[index+1]の重複要素を消す
        ancestor[index+1] = sorted(list(set(ancestor[index+1])))
        index += 1

    # dictionary（特許文献とそれに対する引用文献の関係）をhtml用に整形
    for key in dictionary:
        items = ""
        for ele in dictionary[key]:
            items += ("<li>"+ele+"</li>\n")
        dictionary[key] = key+"\n<ul>\n"+items+"</ul>\n"

    # ツリーを作成
    text = "<!DOCTYPE html>\n<html>\n<ul>\n<li>" + docname + "</li>\n</ul>\n</html>"
    for j in range(1, len(ancestor)):
        for ele in ancestor[j-1]:
            text = text.replace(ele, dictionary[ele])

    # ツリーにurlのリンクを張る
    for j in range(len(ancestor)-1):
        for ele in ancestor[j]:
            literal = url_dictionary[ele]
            text = text.replace(ele, '<a href="' + literal +
                                     '", target=_blanck>' + ele + '</a>')

    # ツリーをファイルに保存
    dirpath = os.path.join(os.getcwd(), '%sOA' % docname)
    with open(os.path.join(dirpath, 'tree.html'), 'w') as f:
        f.write(text)

    # 処理の完了を通知
    alert('Complete')
    sys.exit()
