---
date: 2023-02-08
version: 0.0.0.9000
---

# Text Data Base For VBA について

このプログラムはExcel VBA上で動作する疑似データベース環境を構築します。

例えば、SQLiteではフォルダ上にバイナリファイルを作成し、SQL構文によりデータベース管理を行います。

私が作成したプログラムでは、指定した場所に疑似データベース管理用で `data` フォルダを作成し、このフォルダ内にテーブルを作成していきます。

テーブルはタブ区切りテキストファイル(tsv)形式で個別のファイルとして作成されます。

テーブルの作成・修正等はVBAの `open` メソッドにより行っており、通常環境ではエンコードは `Shift-JIS` となるはずです。

なお、、データ中に改行文字が混じっていることを考慮して、改行文字はWindowsでは基本的に使用しない *VbCr* を使用しています。

このプログラムは、閉鎖的な環境でセキュリティが厳しく、自由にアプリケーションをインストールすることができないオフィス環境で、Excel上のデータを複数人と共有して使用するために作成しました。

VBAによる操作は **SQL** 構文と似ているため、データベースを操作したことがあれば簡単に使うことができます。

# 設定方法

1. 以下のクラスモジュールをVBE画面でインポートしてください。

- textSQL
- HTMLmacro

2. 標準モジュールを追加し、クラスモジュールを使用できるようにグローバル変数として設定しておきます。

```{vba}
Private TSQL As New textSQL
Private HTML As New HTMLmacro
```

3. データベースを作成する場所(path)を設定してください。

```{vba}
TSQL.WorkDir = "folder-full-path"
```

または、**textSQL** クラスモジュールの起動時設定に直接書き込んでください。

```{vba}
Private Sub Class_Initialize()
    working_dir = "folder-full-path" & "\data"
End Sub
```


# 使い方

基本的なSQL操作を実行する `SQL` 関数を準備していますが、引数情報を分かりやすくした `SQLのラッpカー関数` を準備しました。

以下のコードブロックでは、それぞれ2行記載していますが、 `SQL()` はbaseの関数で、2行目はラッパー関数であるため、どちらか一方を実行するとだけでOKです。

## 新規でデータベース登録

以下のコマンドを実行することで、 `WorkDir` で指定したフォルダ内に `data` フォルダが作成されます。

また、バックアップ作成用のフォルダ `Backups` が `data` フォルダ内に作成されます。


```{vba}
TSQL.CREATE_DataBase
TSQL.SQL("CREATE", "DATABASE")
```



## 新規データテーブルの作成

以下の関数でテーブルを作成します。

なお、データ型とデータ列名には1次元配列 `array()` を使って指定します。データ型とデータ列名は必ず同じ長さにしてください。

`primary key` は1列目が自動で指定されます。1列目には重複がないデータを指定してください。

なお、 `runif()` 関数を使用すると、重複なしの数値を発行するので、この番号が入る列を第一列に指定してもOKです。

データ型で設定することができるのは、 **文字列,日付,時間,数値** の4つです。

```{vba}
Call TSQL.SQL("CREATE", "TABLE", "テーブル名", _
        Array("データ型", ...), _
        Array("データ列名", ...))
Call TSQL.CREATE_DataTable("テーブル名", _
	    Array("データ型", ...), _
        Array("データ列名", ...))
```

作成されたテーブルは `\data` 内にテキストファイルとして配置されます。

## データテーブルにデータを追加

テーブルにデータを登録するには、登録するデータを配列かDictionary型で指定する必要があります。

最も簡単な方法は、Excelシート上に列名と登録するデータを準備し、VBAの `Range` で範囲指定をして取り込む方法です。

```{vba}
dlist = Range(Cells(1, 1), Cells(6, 5)) 
Call SQL("UPDATE", "テーブル名", dlist)
Call TSQL.UPDATE_Table("テーブル名", dlist)
```

## データの取得

テーブルから登録しているデータを取得します。

データはDictionary型のオブジェクトとして取得します。

また、SQL構文と同じように `SELECT` に続けで `WHERE, ORDER, JOIN` で操作を続けることができます。

Dictionary型のデータは `dic2arr` 関数で2次元配列に変換されるため、そのままExcelシート上にアウトプットすることができます。

なお、 `SELCT_Table()` 関数では `WHERE, ORDER, JOIN` を引数で1回しか指定できませんが、 `SQL()` 関数では複数回使用できます。

`WHERE` の条件式で使用できる比較演算子は以下の表のとおりです。

| 演算子 | 挙動     |
|--------|----------|
| ==     | 完全一致 |
| =      | 部分一致 |
| !=     | 否定     |
| >      | 大なり   |
| >=,=>  | 以上     |
| <      | 小なり   |
| <=,=<  | 以下     |
| &&     | And検索  |
| \|\|   | OR検索   |



```{vba}
Set dic = SQL("SELECT", "*", "テーブル名")
Set dic = TSQL.SELECT_Table("テーブル名")


'' 他の条件を指定する場合
Set dic = TSQL.SELECT_Table("テーブル名", Array(列名を配列で指定), 絞り込み検索の指定, "検索条件", 並び替えの指定, "並び替えの起点となる列名", テーブル結合の指定, "結合する際のキー列名", "結合するテーブル名")
set dic = TSQL.SQL("SELECT", "*", Array(列名を配列で指定), "WHERE", "検索条件", "ORDER", "並び替えの起点となる列名", 昇順(0)|降順(1), "JOIN", "結合する際のキー列名", "結合するテーブル名")

'' 以下、具体例


'' 列を指定して表示
Set dic = SQL("SELECT", Array("列名1", "列名2", "列名3", ...), "テーブル名")
Set dic = TSQL.SELECT_Table("テーブル名", Array("列名1", "列名2", "列名3", ...))


'' 別のテーブルを結合して表示
Set dic = SQL("SELECT", "*", "テーブル名", _
                "JOIN", "結合する際のキー列名", "結合するテーブル名")
Set dic = TSQL.SELECT_Table("テーブル名", "*", , , , , , True, "結合する際のキー列名", "結合するテーブル名")


'' 検索表示
Set dic = SQL("SELECT", "*", "Asbesto", _
                "WHERE", "列名1>=2023-03-01")
Set dic = TSQL.SELECT_Table("Asbesto", "*", True, "列名1>=2023-03-01")


'' 並び替え処理
Set dic = SQL("SELECT", "*", "テーブル名", _
                "ORDER", "列名1", 0|1)
Set dic = TSQL.SELECT_Table("テーブル名", , False, , True, "列名1", orderASC|orderDESC)


'' 取り出したデータ(Dictionary型)を2次元配列に変換してWorksheetに出力方法
Range(Cells(1, 10), Cells(UBound(data, 1), 10 + UBound(data, 2))) = TSQL.dic2array(data)
```


## データを修正してテーブルを更新

登録したデータを `SELECT` 構文で取り出したあと、修正を行い再度テーブルに登録(上書き)します。

Excelのシート上でデータの修正を行った場合、 `Range` で列名とデータを範囲指定で取得しで処理を行うことができます。(データ型を同じ配列に含める必要はありません。)

```{vba}
data = Range(Cells(1, 1), Cells(6, 5))
Call SQL("UPDATE", "テーブル名", data)
Call TSQL.UPDATE_Table("テーブル名", data)
```

## 新規列の挿入

作成済みのテーブルに、新たに列を追加する場合は `INSERT` 構文を使います。

引数として挿入する列番号を指定しますが、新たな列は指定した列番号の後ろに追加されます。(3と指定した場合、4列目に列が追加)

列の追加には列名と併せてデータ型の指定が必要です。

```{vba}
Call SQL("INSERT", "テーブル名", Array("データ型の指定"), Array("データ列名の指定"), 挿入する列番号(Lng))
Call TSQL.INSERT_Columns("テーブル名", Array("データ型の指定"), Array("データ列名の指定"), 挿入する列番号(Lng))
```

## テーブルからデータを削除

テーブルからデータを削除する場合、 `DELETE` 構文を使います。

削除するデータ名はprimary keyとして登録した列データを指定します。(Dictionary型オブジェクトのkeyになっているデータです。)

テーブル内全てのデータを削除する場合は、列名に **"*"** を指定します。

```{vba}
Call TSQL.SQL("DELETE", "テーブル名", Array("削除データキー名", ...))
Call DELETE_Items("テーブル名", Array("削除データキー名", ...))
```

## 作成したテーブルを削除

テーブルを削除する場合は、 `DROP` 構文を使います。

```{vba}
Call TSQL.SQL("DROP", "テーブル名")
Call TSQL.DROP_Table("テーブル名")
```


## その他の機能

### データベース内にフォルダーを作成

```{vba}
Call TSQL.MakeDirectory("フォルダ名")
```

### テーブル名一覧取得

```{vba}
MsgBox Join(TSQL.tables, VbCrLf))
```

### テーブル内の列名取得

```{vba}
MsgBox Join(TSQL.table_ColNames("テーブル名"))
```

### テーブル内のデータ数取得

```{vba}
MsgBox Join(TSQL.table_Count("テーブル名"))
```

### Dictionary型データを横に結合

```{vba}
Set dic1 = TSQL.SELECT_Table("テーブル名1")
Set dic2 = TSQL.SELECT_Table("テーブル名2")
set TSQL.cbind(dic1, dic2)
```

### Dictionary型データを縦に結合

```{vba}
Set dic1 = TSQL.SELECT_Table("テーブル名1")
Set dic2 = TSQL.SELECT_Table("テーブル名2")
set TSQL.rbind(dic1, dic2)
```

### ランダムな番号の作成

指定した範囲で重複なしの数値を生成します。

作成されるデータは、 **日付_時刻_ランダム値** のようになっています。

```{vba}
num = TSQL.runif(生成するランダム値の数(Lng), 最小値(Lng), 最大値(Lng))
```

### Dictionary型データを2次元配列に変換

```{vba}
Set dic1 = TSQL.SELECT_Table("テーブル名1")
data = TSQL.arr2dic(dic1)
```

### 2次元配列をDictionary型に変換

```{vba}
data = Range(Cells(1, 1), Cells(5, 5))
set dic = TSQL.arr2dic(data)
```

### Dictionary型データからデータを検索取得

`SELECT` 構文中の `WHERE` と同じ働きをします。

```{vba}
Set dic1 = TSQL.SELECT_Table("テーブル名1")
Set Dic2 = TSQL.filter4dic(dic1, "列名1>=100")
```

### Dictionary型データから列を指定してデータを取得

`SELECT` 構文中でテーブルからデータを取得する際、取得する列名を設定することと同じ働きをします。

```{vba}
Set dic1 = TSQL.SELECT_Table("テーブル名1")
Set Dic2 = TSQL.select4dic(dic1, Array("列名", ...))
```

## テーブルのデータをブラウザで表示

ここからは **HTMLmacro** クラスを使用します。

このクラスでは、HTMLを生成・出力するために関数を準備しています。

例えば、 **<HTML>hogehoge</HTML>** を書くには `.HTML("hogehoge")` 関数と書きます。

同じように、 `.h1(), .h2(), h3(), .p(), table(), .article()` 等を使用することができます。

詳しくはクラスモジュールを呼んでみてください。

これらの関数では全て文字列として出力されます。作成された文字列をHTMLファイルに出力し、**Edgeブラウザ** で表示するように `PrintHTMLandOpenByBrowser()` を準備しています。

なお、作成されたHTMLファイルは標準で5秒後に削除される設定になっています。ファイルを残す場合は、`PrintHTMLandOpenByBrowser()` を修正するか、ブラウザに表示されたページを新規で保存してください。

通常のHTMLと同じように、自由にフォーマットを作成することができますが、手っ取り早くデータを確認する場合は、 `TSQL.SELECT_Table()` で取り込んだDictionary型データを、 `.table()` の引数に指定する方法です。

```{vba}
Set dic = TSQL.SELECT_Table("テーブル名")
body = HTML.table(dic, "wide"|"long")
text = HTML.HTML("title", body)
Call HTML.PrintHTMLandOpenByBrowser(text)
```

以下に簡単なサンプルを示します。

```{vba}
Private TSQL As New textSQL
Private HTML As New HTMLmacro
Sub sample(ByVal data_table_name As String)
    
    Dim dic As Object
    Dim body As String
	Dim text as String

    
    '' データテーブル取込み
    Set dic = TSQL.SELECT_Table(data_table_name)
    
    '' 該当するデータを取り出し
    body = ""
    
	'' HTMLのフォーマット指定
    With HTML
        body = body & _
            .article( _
                .h1(.font(data_table_name, "red")) & _
                .table(dic, "long"))
        
    End With
    
    '' HTML出力
    text = HTML.HTML("テスト - HTML出力", body)
    
    Call HTML.PrintHTMLandOpenByBrowser(text)
    
End Sub
```

# 活用方法

私はこのクラスモジュールを簡易的なデータベースとして活用できるよう、ExcelBook上にボタンを配置し、これらの関数を自由に呼び出せるように設定しています。

そのようにすることで、オフィス内の閉鎖的な環境でのみ動作する簡易データベースとすることができます。

これは最初に書いたとおり、セキュリティ上の問題でSQLite等の素晴らしいデータベースを使うことができない場合において、複数人で共有することができるデータベースを構築することができます。

使用が想定される場面はごく少数だと思いますが、参考になれば幸いです。

また、データテーブルはテキストファイルであるため、アプリケーションの管理者が不在時、バグによりVBAが機能しなくなった場合においても、テキストファイルを開くことでデータを確認することができます。

テキストファイルはtsv形式のため、そのままコピー＆ペーストすることで、Excelシート上で綺麗に確認することができます。

