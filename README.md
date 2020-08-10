# はじめに


- この記事は、ITに詳しくないもののPythonで業務を自動化してみたいなと考えている知的財産業務に従事している人を想定しています。

- また、このプログラムを利用して生じたいかなる損失について、責を負いません。

- このプログラムを改変したPatAncestorというプログラムもあります。拒絶理由通知書から主引例を辿って、大元の文献を探すプログラムです。

- このプログラムは[github](https://github.com/inusarukiji/jpo_automation)にもあります。


# 概要


出願人が特許を取得するために必要な書類を特許庁に提出すると、多くの場合、権利範囲を変更せずに特許権を取得することはできません。提出した書類に記載された技術思想に近い先行技術文献が見つかるため、権利範囲を狭める手続きが必要となるからです。この権利範囲を狭める手続きが必要であることを知らせるために、特許庁は出願人に対して拒絶理由通知書を送信します。拒絶理由通知書には、出願人が思いついた技術と似たような技術が記載された先行技術文献（以下、引用文献）が記載されています。引用文献の多くは、特許文献であり[特許庁のサイト](https://www.j-platpat.inpit.go.jp/)にアクセスして閲覧することができます。上記引用文献が閲覧できるurlを取得し、クリックするだけでサイトが開ける拒絶理由通知書のhtmlファイルを生成するプログラムを作成したので公開します。

＊実務上、拒絶理由通知書に記載された引用文献をwebで閲覧することに大した労力は必要ありませんが、本プログラムと似たようなプログラムを作成する際に参考となる情報が含まれていると思います。

# 以下の環境で動作を確認
windows10
python version 3.7.3
Google Chrome version 77.0.3865.75
ChromeDriver version 76.0.3809.126
selenium version 3.141.0
pyperclip version 1.7.0

＊Google ChromeとChromeDriverはversionを揃える必要があるようです。当初は双方ともversion 76.0でしたがChromeは2019/09/12時点でversionが77.0になっています。動作確認したところ動きました。

# プログラムの要点


１．拒絶理由通知書を取得する特許文献をGUIで入力する
２．特許庁のホームページにアクセスし、拒絶理由通知書が登録されていれば取得する
３．拒絶理由通知書にurlを取得できる引用文献が記載されていればurlを取得する
４．クリックすれば引用文献にアクセスできるhtml版の拒絶理由通知書を作成する

# 掲載したプログラムを読み解くと出来るようになること

- PythonでGUIを使用する
- GUIの背景色を設定する
- GUIのフォントを設定する
- GUIに入力された文字列を変数に格納する
- webにプログラムでアクセスする
- webの入力フォーム欄に文字列を入力してサーバに送信する
- web上で所望の`<a>`タグにアクセスする
- 文章内の特定の文字列を変更する
- フォルダを作成してファイルを格納する
- ファイル名を決定してファイルを作成する

# 使用するモジュール、クラス、メソッド等


- selenium
    - webdriver
- tkinter 
    - TK, Label, Entry, Button, StringVar, font
- re
    - search, findall 
- time
    - sleep
- os
    - makedirs, path.join, getcwd
- sys
    - exit
- pyperclip
    - paste

# ソースコード


```python:OAlink.py
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

```

# 各関数についてのコメント


- interface()
interface()関数内でbtn_click()関数を定義しています。これは、Buttonクラスでcommand=btn_clickのオプションを引数に含めるためには、先にbtn_click()関数が定義されていないとエラーが出たため。なお、root.quit()でループを抜ける（window表示を止める）とプロセスが残るのでroot.destroy()でループを抜けています。

- alert(result)
処理の最後でseleniumを終了させるには、driver.close()ではなくてdriver.quit()とします。

- hitnumber(text)
GUIでの入力テキストが不完全だと文献が１つに特定されないため、ここで検索結果を確かめます。

- oagetter()
webページの経過情報ボタンを押して拒絶理由通知書を取得します。ここでは、lst[-1].click()として最後に発行された拒絶理由通知書を取得する仕様となっています。

- filecreator(OAText, dnum1, dirpath)
カレントディレクトリ（ターミナルまたはコマンドプロンプトが動いているところ）内にフォルダを作成し、そのフォルダの中に拒絶理由通知書のhtmlファイルを格納します。

- doc_catcher(OAText)
まず、拒絶理由通知書の中から範囲を絞ります。特許出願の番号と先行技術文献調査結果に記載されている文献番号を拾わないようにするためです。

- jp_doc_linker(p_cited, filepath)
引用文献のうち日本語の特許文献のリンクを取得します。併せて、拒絶理由通知書のhtmlファイルをリンク付きのものに書き換えます。

- us_doc_linker(usp_cited, filepath)
引用文献のうち外国語の特許文献のリンクを取得します。ここでは、例として米国の特許文献のリンクを取得する仕様となっています。併せて、拒絶理由通知書のhtmlファイルをリンク付きのものに書き換えます。

# クローラーを作る上で参考にしたサイト


- @nezuq さんの[記事](https://qiita.com/nezuq/items/c5e827e1827e7cb29011)。web上でbotを走らせる際の注意点が説明されている
- @Azunyan1111 さんの[記事](https://qiita.com/Azunyan1111/items/b161b998790b1db2ff7a) 。webスクレイピングのテクニックがまとめられている
- [MDNのサイト](https://developer.mozilla.org/ja/docs/Learn/Getting_started_with_the_web)。html, css, javascript, httpについて学べる。
- [seleniumのクイックリファレンス](https://www.seleniumqref.com/)。逆引き辞典があって便利。
- [Xpathのチートシート](http://aoproj.web.fc2.com/xpath/XPath_cheatsheets_v2.pdf)。関数がたくさん載っている。
- Tkinterの背景色の一覧が載っている[サイト](http://memopy.hatenadiary.jp/entry/2017/06/11/092554)。stackoverflowにすべて？の色を表示する[プログラム](https://stackoverflow.com/questions/4969543/colour-chart-for-tkinter-and-tix)が公開されている模様。

# 自分用備忘録


- 当初はhtmlファイルで、別のwindowBを開きwindowAからjavascriptを実行することによってwindowBを操作するというプログラムを作ろうと思った。しかしwebブラウザのセキュリティ上の制限により上の動作は実行できなかった。具体的には、

```javascript:html
<script>
...
subwin = window.open('url');
subwin.document.getElementById('IDname');
...
</script>
```
としても、以下の例外メッセージがでた。

> Uncaught DOMException: Blocked a frame with origin "file://" from accessing a cross-origin frame.

上記制限は、[CSRF(cross-site request forgeries)](https://ja.wikipedia.org/wiki/%E3%82%AF%E3%83%AD%E3%82%B9%E3%82%B5%E3%82%A4%E3%83%88%E3%83%AA%E3%82%AF%E3%82%A8%E3%82%B9%E3%83%88%E3%83%95%E3%82%A9%E3%83%BC%E3%82%B8%E3%82%A7%E3%83%AA)といった脅威を回避するために設けられているそう。また、Google Chromeの場合、ショートカットのリンク先を"....exe"  --disable-web-security --user-data-dir="C://Chrome dev session"と変更して起動することにより、このwebブラウザの制限は取り除かれる（脅威に関して脆弱になる）ようだが、試しても上手くいかなかった。
なお、ローカルの同一ディレクトリにhtmlファイルAとhtmlファイルBを置いて、ブラウザ上でjavascriptによりファイルBをファイルAから操作することはできるようになった。また、internet explorerとEdgeでは、何の変更もなしに同じことができた。しかしながら、internet explorer, Edge, Firefox, Chromeのいずれでもhtmlファイルからwebサイトを開き、htmlファイル側からそのサイトのdocumentを読み込むことはできなかった。

- 上で述べたwebブラウザの制限は、[CORS(Cross Origin Resource Sharing)](https://developer.mozilla.org/ja/docs/Web/HTTP/CORS)や[JSONP(Javascript Object Notation with Padding)](https://ja.wikipedia.org/wiki/JSONP)を利用すると解決できるようだが、JSONPについては[通信内容が秘匿されない](https://blog.ohgaki.net/stop-using-jsonp)ため慎重になった方がいいみたい。

- seleniumでdriver.switch_to.window(...)をした直後にDOMのelementをclick( )しようとすると以下のような例外メッセージがでた。

> ElementClickInterceptedException: Message: element click intercepted: Other element would receive the click

このメッセージを見て、取得しているelementが違うのかとか、javascriptでelementの属性が動的に変化していて無効なのかとか色々検討したがswitch_to.windowの後にtime.sleep()を入れたら解決した。

- webページ中、どのelementを操作しているか把握するためにChromeの拡張機能であるselenium IDEとKatalon Recorderを利用してみた。Katalon Recorderは、VBAのマクロ機能のようにweb操作を記録することができ、記録した内容を各種プログラミング言語に変換する機能を提供する（selenium IDEも以前はできたが2019/09/13時点で休止中）。しかしながら、操作するelementのXpathが、あるelementから５つ目の\<div>といった形で表現されていてあまり参考にならなかった(javascriptで動的にelementの属性が変化するのでそうしているのだろうけど)。通常どおり、chromeのweb開発ツールを参照するのがよい。

- htmlのelementの指定には、複数のやり方があるが、Xpathの関数による指定が一番きめ細かくelementを指定できた。

- seleniumのfind_element_by_idは単一のelementを取得するためclick()やsend_keys()などの操作ができるが、find_elements_by_idはリストを返すため同じ操作はできない。

- 検索して他人のブログや質問サイトの回答を覗くより、初めからリファレンスやチュートリアルを読んだ方が早く解決手段が見つかった。

- PythonのrequestsモジュールではJavascriptで動的に変化するサイトはスクレイピングが難しいのかと思ってseleniumを使った。しかし、http通信について理解すればrequestsモジュールでも同じものが作れそう。その場合、例えば、Chromeのweb開発ツールのnetworkタブを参照する。







### Patancestor.py
１回目の拒絶理由通知の主引例（引用文献１）を親とし、親の親、親の親の親・・・をたどる。
