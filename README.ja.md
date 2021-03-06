# X2P-QuizMaker

Version 1.0.2

![デモ](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/demo.gif)

[English README](README.md)


## 目次

- [概要](#%e6%a6%82%e8%a6%81)
- [特徴](#%e7%89%b9%e5%be%b4)
- [用語](#%e7%94%a8%e8%aa%9e)
- [使い方](#%e4%bd%bf%e3%81%84%e6%96%b9)
    - [事前準備](#%e4%ba%8b%e5%89%8d%e6%ba%96%e5%82%99)
    - [実行](#%e5%ae%9f%e8%a1%8c)
    - [実行結果](#%e5%ae%9f%e8%a1%8c%e7%b5%90%e6%9e%9c)
- [ひな型ファイルの作り方 (PowerPoint)](#%e3%81%b2%e3%81%aa%e5%9e%8b%e3%83%95%e3%82%a1%e3%82%a4%e3%83%ab%e3%81%ae%e4%bd%9c%e3%82%8a%e6%96%b9-powerpoint)
    - [OK なこと](#ok-%e3%81%aa%e3%81%93%e3%81%a8)
- [クイズリストファイルの作り方 (Excel)](#%e3%82%af%e3%82%a4%e3%82%ba%e3%83%aa%e3%82%b9%e3%83%88%e3%83%95%e3%82%a1%e3%82%a4%e3%83%ab%e3%81%ae%e4%bd%9c%e3%82%8a%e6%96%b9-excel)
    - [ひな型ファイルのパスを貼り付ける](#%e3%81%b2%e3%81%aa%e5%9e%8b%e3%83%95%e3%82%a1%e3%82%a4%e3%83%ab%e3%81%ae%e3%83%91%e3%82%b9%e3%82%92%e8%b2%bc%e3%82%8a%e4%bb%98%e3%81%91%e3%82%8b)
    - [リストを作成する](#%e3%83%aa%e3%82%b9%e3%83%88%e3%82%92%e4%bd%9c%e6%88%90%e3%81%99%e3%82%8b)
    - [マクロを追加する](#%e3%83%9e%e3%82%af%e3%83%ad%e3%82%92%e8%bf%bd%e5%8a%a0%e3%81%99%e3%82%8b)
    - [OK なこと](#ok-%e3%81%aa%e3%81%93%e3%81%a8-1)
- [License](#license)


## 概要

Excel ワークシートに入力したクイズリストを PowerPoint のスライドに流し込むマクロです。


## 特徴

- PowerPoint のスライドは、デザインやレイアウト、アニメーションを自由に設定可能
- 右クリックからマクロを実行可能


## 用語

<dl>
  <dt>ひな型ファイル</dt>
  <dd>Excel に入力したクイズの貼り付け先となる PowerPoint ファイル。<br>このファイルをコピーし、そこにクイズを貼り付ける。</dd>
  <dt>ひな型スライド</dt>
  <dd>ひな型ファイルの中の、実際にクイズを貼り付けるスライド。</dd>
  <dt>クイズリストファイル</dt>
  <dd>クイズリストを入力した Excel ファイル。</dd>
  <dt>クイズリスト</dt>
  <dd>クイズリストファイルの内容。</dd>
</dl>


## 使い方


### 事前準備

1. 「[ひな型ファイルの作り方 (PowerPoint)](#%e3%81%b2%e3%81%aa%e5%9e%8b%e3%83%95%e3%82%a1%e3%82%a4%e3%83%ab%e3%81%ae%e4%bd%9c%e3%82%8a%e6%96%b9-powerpoint)」を参考に、PowerPoint でひな型ファイルを作成してください
1. 「[クイズリストファイルの作り方 (Excel)](#%e3%82%af%e3%82%a4%e3%82%ba%e3%83%aa%e3%82%b9%e3%83%88%e3%83%95%e3%82%a1%e3%82%a4%e3%83%ab%e3%81%ae%e4%bd%9c%e3%82%8a%e6%96%b9-excel)」を参考に、Excel でクイズリストファイルを作成してください


### 実行

1. クイズリストファイルを開いてください
1. クイズリストファイルのシート上で、適当なセルを右クリックしてください
1. メニューの一番下あたりにある、クイズリストファイルと同じ名前の項目をクリックしてください
1. 右に飛び出した項目の中の「クイズをスライドに流し込む(Q)」をクリックしてください（マクロが実行されます）


### 実行結果

クイズリストの各行の項目が、ひな型スライドに貼り付けされます。

![クイズリスト](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_excel_quiz_1.png)

![ひな型スライド](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_ppt_1.png)

![実行結果](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_result.png)

使用するひな型スライドは、クイズリストの `A` 列で指定します。上記のクイズリストの "Quiz 1" の場合、`A4` セルに 2 と入力してありますので、2枚目のひな型スライドが使用されています。

クイズリストの `{title}` の内容 "Quiz 1" が、ひな型スライドの `{title}` の箇所に貼り付けされています。同様にクイズリストの `{content}` の内容 "10進法の 10 を2進法に変換したものはどれ？" がスライドの `{content}` に、クイズリストの `{1}` の "10" がスライドの `{1}` に、それぞれ貼り付けされています。


## ひな型ファイルの作り方 (PowerPoint)

![ひな型ファイル](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_ppt_1.png)

- クイズや選択肢を貼り付けるテキストボックスや図形には、目印となる文字列を入力してください  
    クイズの文章を貼り付ける箇所には `{content}`、1番の選択肢を貼り付ける箇所には `{1}` などです。  
    クイズリストファイルの3行目と同じであれば、どんな文字列でも構いません
- ひな型**スライド**は、選択肢の数だけ作成してください  
    4択クイズの場合は、1番が正解の時のスライド、2番が正解の時のスライド、3番のスライド、4番のスライド、という風に4枚必要です  
    ![ひな型スライドは選択肢の数だけ作成](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_ppt_1_slides_for_answers.png)
- ひな型スライドにアニメーションを設定してください  
    1番のスライドには1番が正解の時のアニメーション、2番の時のスライドには2番が正解のアニメーション、という風に、**それぞれのひな型スライドに対して設定してください**
- ひな型スライドは非表示にしてください  
    ![ひな型スライドは非表示](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_ppt_1_invisible.png)
- 表紙のスライド等は、必要に応じて作成してください  
    ![表紙のスライド](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_ppt_1b_title.png)


### OK なこと

- ファイル名の変更  
    ただし、クイズリストファイルの `A1` セルも変更してください
- ひな型スライドの追加や削除  
    選択肢の数だけ作成してください
- スライドのレイアウトの変更  
    タイトルの有無、選択肢の位置や数、デザインなど、自由に設定してください
- アニメーションの変更


## クイズリストファイルの作り方 (Excel)

![クイズリストファイル](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_excel.png)


### ひな型ファイルのパスを貼り付ける

- ひな型ファイルのパスを `A1` セルに貼り付けてください  
    フォルダー上で、`Shift` キーを押したままひな型ファイルを右クリックし、「パスのコピー(A)」をクリックするとコピーできます。そのまま `A1` セルに貼り付けてください  


### リストを作成する

- `A3` セル以降にクイズリストを作成してください
- `A` 列には、そのクイズに対して使用するスライド番号を入力してください  
- クイズリスト内に空白の行や列を作らないでください
- クイズリストに隣接するセルには、何も入力せず、空けておいてください  
    ![クイズリストに隣接するセル](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_excel_keep_blank.png)
- このファイルを「Excel マクロ有効ブック (*.xlsm)」として保存してください


### マクロを追加する

以下の記事を参考にしてください。  
https://taidalog.hatenablog.com/entry/2022/05/05/100000

1. クイズリストファイルを開いてください
1. [Releases · taidalog/X2P-QuizMaker](https://github.com/taidalog/X2P-QuizMaker/releases) から最新版の「Source code (zip)」をダウンロードしてください
1. ダウンロードした ZIP ファイルを展開（解凍）してください
1. 展開（解凍）して手に入れた `main.bas` をクイズリストファイルにインポートしてください
1. クイズリストファイルの `ThisWorkbook` モジュールに以下のコードをコピーして貼り付けてください  
    ```
    Private Sub Workbook_Open()
        Call AddToContextMenu
    End Sub
    ```
1. クイズリストファイルを保存して閉じて、再度開いてください  
    適当なセルを右クリックして、コンテキストメニュー（右クリックメニュー）の一番下辺りに、クイズリストファイルと同じ名前の項目があればインポート成功です


### OK なこと

- ファイル名やシート名の変更
- シートの追加  
    シートごとに、異なるひな型ファイルを指定できます
- クイズリストの項目の変更・追加・削除  
    `{title}`, `{content}`, `{1}`, `{2}` ... の部分は変更可能です。ただし、それに合わせてひな型スライドの方も変更してください
- 選択肢の数の増減  
    それに合わせてひな型スライドの方も変更してください  
    ![クイズリスト（3択）](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_excel_2.png)
    ![ひな型ファイル（3択）](https://github.com/taidalog/X2P-QuizMaker/blob/images/image/X2P_ppt_2b.png)
- クイズリストの行数（＝クイズの設問数）の増減  
    Excel そのものの行数と、マクロの実行時間以外に制限はありません


## License

Copyright 2022 taidalog

X2P-QuizMaker is licensed under the MIT License.
