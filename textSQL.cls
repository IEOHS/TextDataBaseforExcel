VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "textSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private working_dir As String
Public Enum readtype
    RowHead = 0
    Number = 1
End Enum
Public Enum SortOrder
    orderASC = 0
    orderDESC = 1
End Enum
Public Enum SepKey
    keyComma = 1
    keyTab
End Enum
Public Function runif(ByVal get_number As Long, ByVal nmin As Long, ByVal nmax As Long) As Variant
    '' ランダム数値の作成
    '' get_number: 取得する数値の数
    '' nmin: 最小値
    '' nmax: 最大値
    
    Dim n As Long
    Dim dic As Object
    
    '' ランダム値の生成
    Set dic = Set_Dictionary
    '' シードの初期化
    Randomize
    '' 日付_時間_ランダム数値として作成
    Do While dic.Count <> get_number
        n = Int((nmax - nmin + 1) * Rnd + nmin)
        If dic.exists(n) = False Then
            dic.Add n, Format(Now(), "yyyyMMdd_HHmmSS_") & Format(n, "00000")
        End If
    Loop
    
    runif = dic.items
    
End Function

Public Function CREATE_DataBase()
    Call Set_New_Data_Base
End Function
Public Function CREATE_DataTable(ByVal create_table_name As String, ByVal table_column_types As Variant, _
                                ByVal table_column_names As Variant)
    Call SQL("CREATE", "TABLE", create_table_name, table_column_types, table_column_names)
End Function
Public Function INSERT_Columns(ByVal table_name As String, ByVal table_column_types As Variant, _
                        ByVal table_column_names As Variant, ByVal set_insert_columns_number As Long) As Object
    '' INSERT table名 データ型配列、データ名配列、挿入位置の数字
    Set INSERT_Columns = Insert_New_Column_To_DataTable(table_name, set_insert_columns_number, _
                                            table_column_types, table_column_names)
            
End Function
Public Function UPDATE_Table(ByVal table_name As String, ByVal update_data As Variant)
    '' UPDATE table名 dictionary名|配列名
    Call SQL("UPDATE", table_name, update_data)
End Function
Public Function DELETE_Items(ByVal table_name As String, ByVal delete_key As Variant)
    '' DELETE table名 *
    '' DELETE table名 削除データキー配列
    Call SQL("DELETE", table_name, delete_key)
End Function
Public Function DROP_Table(ByVal table_name As String)
    Call SQL("DROP", table_name)
End Function
Public Function tables() As Variant
    tables = Get_Table_List
End Function
Public Function table_ColNames(ByVal table_name As String) As Variant
    Dim dic As Object
    
    Set dic = Read_DataTable(table_name)
    table_ColNames = dic.item("name")
End Function
Public Function table_Count(ByVal table_name As String) As Variant
    Dim dic As Object
    Set dic = Read_DataTable(table_name)
    table_Count = dic.Count - 3
End Function
Public Function SELECT_Table(ByVal table_name As String, _
                            Optional view_column_names As Variant = "*", _
                            Optional select_column_names As Boolean = False, _
                            Optional formula As String = "", _
                            Optional order_table As Boolean = False, _
                            Optional names_to_order As String = "", _
                            Optional ASC_or_DESC As SortOrder = SortOrder.orderASC, _
                            Optional table_join As Boolean = False, _
                            Optional join_key_name As String = "", _
                            Optional join_table_name As String = "") As Object
    '' SELECT * table名 WHERE 条件式
    '' SELECT 列名配列 table名 ORDER 列名 ASC|DESC
    '' SELECT 列名配列 table名 JOIN 列名 table名
    Dim connection As Variant
    
    '' 条件式による分岐
    connection = Array(select_column_names, order_table, table_join)
    Select Case join(connection, ",")
        Case join(Array(False, False, False), ",")
            Set SELECT_Table = SQL("SELECT", view_column_names, table_name)
        
        Case join(Array(True, False, False), ",")
            Set SELECT_Table = SQL("SELECT", view_column_names, table_name, _
                                "WHERE", formula)
            
        Case join(Array(False, True, False), ",")
            Set SELECT_Table = SQL("SELECT", view_column_names, table_name, _
                                    "ORDER", names_to_order, ASC_or_DESC)
                                    
        Case join(Array(False, False, True), ",")
            Set SELECT_Table = SQL("SELECT", view_column_names, table_name, _
                                    "JOIN", join_key_name, join_table_name)
            
        Case join(Array(True, True, False), ",")
            Set SELECT_Table = SQL("SELECT", view_column_names, table_name, _
                                    "WHERE", formula, _
                                    "ORDER", names_to_order, ASC_or_DESC)
                                    
        Case join(Array(True, False, True), ",")
            Set SELECT_Table = SQL("SELECT", view_column_names, table_name, _
                                    "WHERE", formula, _
                                    "JOIN", join_key_name, join_table_name)
        
        Case join(Array(False, True, True), ",")
            Set SELECT_Table = SQL("SELECT", view_column_names, table_name, _
                                    "ORDER", names_to_order, ASC_or_DESC, _
                                    "JOIN", join_key_name, join_table_name)
        
        Case join(Array(True, True, True), ",")
            Set SELECT_Table = SQL("SELECT", view_column_names, table_name, _
                                    "WHERE", formula, _
                                    "ORDER", names_to_order, ASC_or_DESC, _
                                    "JOIN", join_key_name, join_table_name)
        
    End Select
    
    
    
End Function
Public Function arr2dic(ByVal table_array As Variant) As Variant
    
    Dim dt() As Variant
    Dim stext As String
    Dim i, j As Long
    
    If InStr("文字列,数値,日付,時間", table_array(LBound(table_array, 1), LBound(table_array, 2))) > 0 Then
        ReDim dt(LBound(table_array, 1) To UBound(table_array, 1) - 1, _
                LBound(table_array, 2) To UBound(table_array, 2))
        For i = LBound(table_array, 1) + 1 To UBound(table_array, 1)
            For j = LBound(dt, 2) To UBound(dt, 2)
                dt(i - 1, j) = table_array(i, j)
            Next
        Next
    End If
    
    Set table_array = Array_to_Dictionary(dt, 1)
    
End Function
Public Function dic2arr(ByVal table_object As Object) As Variant
    dic2arr = Dictionary_to_Array(table_object)
End Function
Public Function SQL(ParamArray sql_code()) '' saaa
    '' SQL構文を使って操作
    
    Dim scode As String
    Dim dic As Object
    Dim i As Long
    Dim ikey As Variant
    Dim filename As String
    Dim ch As String
    
    Select Case sql_code(0)
        Case "CREATE"
            '' CREATE DATATABLE
            '' CREATE TABLE table名 データ型配列 データ列名配列
            
            '' 文字列変換
            If sql_code(1) = "DATABASE" Then
                Call Set_New_Data_Base

            ElseIf sql_code(1) = "TABLE" Then
                Call Add_Data_Table(sql_code(2), 1, sql_code(3), sql_code(4))
                MsgBox sql_code(2) & "を作成しました", vbInformation
                
            Else
                MsgBox "書き方が間違っています。"
                Exit Function
            End If
            
        Case "SELECT"
            '' SELECT * table名 WHERE 条件式
            '' SELECT 列名配列 table名 ORDER 列名 ASC|DESC
            '' SELECT 列名配列 table名 JOIN 列名 table名
            
            If Search_Dimentions(sql_code(1)) = 1 Then
                Set dic = Select_Columns(Read_DataTable(sql_code(2)), sql_code(1))
            ElseIf sql_code(1) = "*" Then
                Set dic = Read_DataTable(sql_code(2))
            Else
                MsgBox "構文が間違ってきます。"
                Exit Function
            End If
            
            If UBound(sql_code) < 3 Then GoTo Continue_Select
            
            For i = 3 To UBound(sql_code)
                Select Case sql_code(i)
                    Case "WHERE"
                        Set dic = SELECT_MODE(dic, sql_code(i), sql_code(i + 1))
                        i = i + 1
                    Case "ORDER", "JOIN"
                        Set dic = SELECT_MODE(dic, sql_code(i), sql_code(i + 1), sql_code(i + 2))
                        i = i + 2
                End Select
            Next
            
Continue_Select:
            
            Set SQL = dic
            
        Case "INSERT"
            '' INSERT table名 データ型配列、データ名配列、挿入位置の数字
            Set SQL = Insert_New_Column_To_DataTable(sql_code(1), sql_code(4), sql_code(2), sql_code(3))
            
        Case "UPDATE"
            '' UPDATE table名 dictionary名|配列名
            Call Write_DataTable(sql_code(1), Update_DataTable(sql_code(1), sql_code(2)))
            Call Make_Backup_File(sql_code(1))
            
        Case "DELETE"
            '' DELETE table名 *
            '' DELETE table名 削除データキー配列
            
            Set dic = Read_DataTable(sql_code(1))
            If sql_code(2) = "*" Then
                For Each ikey In dic
                    If ikey = "primary_key" Or ikey = "name" Or ikey = "type" Then
                    
                    Else
                        dic.Remove ikey
                    End If
                Next
            Else
                If Search_Dimentions(sql_code(2)) = 0 Then sql_code(2) = Array(sql_code(2))
                For Each ikey In sql_code(2)
                    dic.Remove ikey
                Next
            End If
            ''Set SQL = dic
            Call Write_DataTable(sql_code(1), dic)
            Call Make_Backup_File(sql_code(1))
        
        Case "DROP"
            '' DROP table名
            filename = DT_name(sql_code(1))
            ch = MsgBox(filename & vbCrLf & "このテーブルを削除しますか?", vbYesNo)
            If ch = 6 Then
                Kill filename
                Debug.Print "Remove File: " & filename
                MsgBox filename & vbCrLf & "テーブルを削除しました。", vbInformation
            End If
            
        Case ".table"
            '' .table
            SQL = Get_Table_List
        
        Case ".table.names"
            Set dic = Read_DataTable(sql_code(1))
            SQL = dic.item("name")
    End Select

End Function
Private Function SELECT_MODE(ByVal dic As Object, ByVal mode As Variant, ParamArray p()) As Object
    On Error GoTo Continue
    Dim num As Long
    
    Select Case mode
        Case "WHERE"
            Set SELECT_MODE = Filter_Data_Table(dic, p(0))
        
        Case "ORDER"
'            If InStr(p(1), "ASC") > 0 Then
'                num = 0
'            ElseIf InStr(p(1), "DESC") > 0 Then
'                num = 1
'            Else
'                num = 0
'            End If
'
'            Set SELECT_MODE = Sort_Data_Table(dic, p(0), num)
            Set SELECT_MODE = Sort_Data_Table(dic, p(0), p(1))
            
        Case "JOIN"
            Set SELECT_MODE = Left_Join(dic, p(0), Read_DataTable(p(1)))
            
        Case Else
            Set SELECT_MODE = dic
                
    End Select
    
Continue:
    
    
End Function
Private Sub Make_Backup_File(ByVal DataTable_name As String)
    '' バックアップファイルを作成
    Dim bf As String
    
    bf = Join_Path(working_dir, "Backups\" & Format(Now(), "yyyyMMdd_HHmmSS") & "_backup_file_" & DataTable_name & ".txt")
    FileCopy DT_name(DataTable_name), bf
    
End Sub
Private Function Remove_Rows_From_2dimArray(ByVal dt As Variant, ByVal row_num As Long) As Variant
    '' 2次元配列の対象列を除外
    '' row_num: 除外する行番号
    
    Dim New_dt() As Variant
    Dim i, j As Long
    
    ReDim New_dt(LBound(dt, 1) To UBound(dt, 1) - 1, LBound(dt, 2) To UBound(dt, 2))
    
    For i = LBound(dt, 1) To UBound(dt, 1)
        For j = LBound(dt, 2) To UBound(dt, 2)
            If i < row_num Then
                New_dt(i, j) = dt(i, j)
            ElseIf i > row_num Then
                New_dt(i - 1, j) = dt(i, j)
            End If
        Next
    Next
    
    Remove_Rows_From_2dimArray = New_dt
    
End Function
Public Function arr2DicArr(ByVal data_array As Variant) As Variant
    arr2DicArr = Get_Table_To_Array(data_array)
End Function
Private Function Get_Table_To_Array(ByVal dt As Variant) As Variant
    '' 二次元配列を列名でdictionaryに登録し、それを配列にする
    '' Update_DataTableのインプットデータに使用
    '' dt: ２次元配列
    '' 最初の行は列名にする
    '' | 列名　| 列名　| 列名...|
    '' | 値　　| 値   | 値... |
        
    Dim New_data() As Variant
    Dim col_name() As Variant
    Dim i, j, num As Long
    Dim bCell As Variant
    
    ReDim New_data(LBound(dt, 1) To UBound(dt, 1) - 1)
    ReDim col_name(LBound(dt, 2) To UBound(dt, 2)) As Variant
    
    '' 列名取得
    For i = LBound(dt, 2) To UBound(dt, 2)
        col_name(i) = dt(LBound(dt, 1), i)
    Next
    
    '' データを登録
    For i = LBound(dt, 1) To UBound(dt, 1) - 1
        
        Set New_data(i) = Set_Dictionary
        For j = LBound(col_name) To UBound(col_name)
            New_data(i).Add col_name(j), dt(i + 1, j)
        Next
        
    Next
    
    Get_Table_To_Array = New_data
    
End Function
Private Function Get_DataTable_Format(ByVal DataTable_name As String) As Object
    '' データテーブルに登録するために、空のdictionaryを取得
    
    Dim n As Integer
    Dim i As LongPtr
    Dim text As String
    Dim dic As Object
    
    Set dic = Set_Dictionary
    n = FreeFile
    
    Open DT_name(DataTable_name) For Input As #n
        i = 1
        Do While Not EOF(n)
            Line Input #n, text
            
            If text = "" Then GoTo Continue
            If i = 3 Then
                For Each ikey In Split(text, vbTab)
                    dic.Add ikey, ""
                Next
                dic.Remove ("NAME")
            End If
            
Continue:
            i = i + 1
        Loop
    
    Close #n
    
    
    Set Get_DataTable_Format = dic
    
End Function
Public Function filter4dic(ByVal data_object As Object, ByVal formula As String) As Object
    Set filter4dic = Filter_Data_Table(data_object, formula)
End Function

Private Function Filter_Data_Table(ByVal data As Object, ByVal formula As String) As Object
    '' Dictionaryタイプで取り込んだall_dataをフィルター処理

    
    Dim d_Type, d_name As Variant
    Dim fs As Variant
    Dim i As Long
    Dim ops As Variant
    Dim base_data, f_data, or_data As Object
    Dim sval, tval, ikey As Variant
    
    ''d_Type = data.item("type")
    d_name = data.item("name")
    
    '' 式と条件を分割
    fs = Separate_Formula(formula)
    
    '' 分岐前のコピーを作成
    Set base_data = data
    
    i = 1
    Do
        '' 比較演算子の判定
        ops = Select_Operator(fs(1)(i), d_name)
        
        '' 検索
        Set f_data = Set_Dictionary
        
        '' データ型の取得
        If ops(0) <> "NA" Then
            d_Type = data.item("type")(ops(1))
        Else
            d_Type = "文字列"
        End If
        
        '' 検索条件値のフォーマット変換
        sval = Text_Format(ops(2), d_Type)
        
        If IsEmpty(sval) Then
            GoTo Continue
        End If
        
        For Each ikey In base_data.keys
            '' 列名とデータ型は必ず登録
            If ikey = "primary_key" Or ikey = "name" Or ikey = "type" Then
                f_data.Add ikey, base_data(ikey)
                GoTo Continue
            End If
            
            '' 対象データを取得
            If ops(0) <> "NA" Then
                tval = Text_Format(base_data(ikey)(ops(1)), d_Type)
            Else
                tval = join(base_data(ikey), ",")
            End If
                
            If IsEmpty(tval) Then GoTo Continue
            
            '' 該当するデータを登録
            Select Case ops(0)
                Case "="
                    If InStr(tval, sval) > 0 Then
                        f_data.Add ikey, base_data(ikey)
                    End If
                    
                Case "=="
                    If tval = sval Then
                        f_data.Add ikey, base_data(ikey)
                    End If
                    
                Case "!="
                    If tval <> sval Then
                        f_data.Add ikey, base_data(ikey)
                    End If
                    
                Case ">=", "=>"
                    If tval >= sval Then
                        f_data.Add ikey, base_data(ikey)
                    End If
                    
                Case ">"
                    If tval > sval Then
                        f_data.Add ikey, base_data(ikey)
                    End If
                    
                Case "<=", "=<"
                    If tval <= sval Then
                        f_data.Add ikey, base_data(ikey)
                    End If
                    
                Case "<"
                    If tval < sval Then
                        f_data.Add ikey, base_data(ikey)
                    End If
                    
                Case "NA"
                    If InStr(tval, sval) > 0 Then
                        f_data.Add ikey, base_data(ikey)
                    End If
                Case "Error"
                    Exit For
                
            End Select

Continue:
        Next
        
        
        '' 前回条件がorの場合の処理
        If i > 1 Then
            If fs(0)(i - 1) = "or" Then
                '' or_data(前回検索したデータ)をf_data(今回検索したデータ)に統合
                For Each ikey In or_data.keys
                    If f_data.exists(ikey) = False Then
                        f_data.Add ikey, or_data(ikey)
                    End If
                Next
                
            End If
        End If
        
        '' 今回条件に対する処理
        Select Case fs(0)(i)
            Case "end"
                Exit Do
            
            Case "and"
                Set base_data = Nothing
                Set base_data = f_data
                Set f_data = Nothing
                
            Case "or"
                Set or_data = Nothing
                Set or_data = f_data
                Set f_data = Nothing
                
                
        End Select

        i = i + 1
    Loop
    
    
    Set Filter_Data_Table = f_data
    

End Function
Private Function Select_Operator(ByVal formula As String, ByVal names As Variant) As Variant
    '' 比較演算子の判定 names 配列型の名前リスト
    '' ==: 完全一致
    '' = : 部分一致
    '' !=: 否定
    '' > : 大なり
    '' >=: 以上
    '' < : 小なり
    '' <=: 以下
    
    Dim operator, str_item As String
    Dim sep_str As Variant
    Dim i, num As Long
    Dim search_text As String
    Dim name_is_left_hand As Boolean
    On Error Resume Next
    
    name_is_left_hand = True
    num = 999
    
    If InStr(formula, "=") > 0 Then
        For Each operator In Array("==", "!=", ">=", "=>", "<=", "=<", "=")
            If InStr(formula, operator) > 0 Then Exit For
        Next
        
    ElseIf InStr(formula, ">") > 0 Then
        operator = ">"
    ElseIf InStr(formula, "<") > 0 Then
        operator = "<"
    Else
        operator = "NA"
    End If
    
    '' 検索文字を分割
    If operator <> "NA" Then
        sep_str = Split(Replace(Replace(formula, " ", ""), _
                                "　", ""), operator)
        
        '' 列番号を検索
        i = 0
        Do
            num = WorksheetFunction.Match(sep_str(i), names, 0) - 1
            
            '' ヒットしなければ右辺を検索
            If num = 999 Then
                name_is_left_hand = False
                If i = 0 Then
                    i = i + 1
                ElseIf i = 1 Then
                    Exit Do
                End If
            Else
                Exit Do
            End If
            
        Loop
        
        '' 右辺でヒットした場合は、比較演算子を逆転
        If name_is_left_hand = False Then
            Select Case operator
                Case "<"
                    operator = ">"
                
                Case "<=", "=<"
                    operator = ">="
                
                Case ">=", "=>"
                    operator = "<="
            End Select
            search_text = sep_str(0)
        ElseIf name_is_left_hand = True Then
            search_text = sep_str(1)
        
        End If
        
        If num = 999 Then
            '' 列名が見つからなかった場合はエラー表示
            operator = "Error"
            search_text = "Error"
       
        End If
        
         
    Else
        search_text = formula
        num = 999
    End If
    
    
    
    Select_Operator = Array(operator, num, search_text)
    
End Function
Private Function Separate_Formula(ByVal formula As String) As Variant
    '' フィルター処理で使う評価式の判定
    '' and検索: &&
    '' or 検索: ||
    
    Dim andLen, orLen, cutStr As Long
    Dim pat, texts As Collection
    Dim formula_text As String
    
    Set pat = New Collection
    Set texts = New Collection
    
    
    '' 空白文字の削除
    formula_text = Replace(formula, " ", "")
    formula_text = Replace(formula_text, "　", "")
            
    '' &&または||で文字が分割できなくなるまで継続
    Do
        '' 条件判定
        andLen = InStr(formula_text, "&&")
        orLen = InStr(formula_text, "||")
        
        '' 条件を振り分け
        If andLen = 0 And orLen = 0 Then
            pat.Add "end"
            texts.Add formula_text
            Exit Do
            
        ElseIf andLen > 0 And orLen = 0 Then
            pat.Add "and"
            cutStr = andLen
            
        ElseIf andLen = 0 And orLen > 0 Then
            pat.Add "or"
            cutStr = orLen
            
        ElseIf andLen > 0 And orLen > 0 Then
            If andLen < orLen Then
                pat.Add "and"
                cutStr = andLen
            ElseIf andLen > orLen Then
                pat.Add "or"
                cutStr = orLen
            End If
            
        End If
        
        '' 文字を分割して登録
        texts.Add left(formula_text, cutStr - 1)
        formula_text = mid(formula_text, cutStr + 2)
        
    Loop
    
    Separate_Formula = Array(pat, texts)
    
    

End Function
Private Function Sort_Data_Table(ByVal dic As Object, ByVal col_name As String, _
                        ByVal order As Long) As Object
    '' dic: dictionary型オブジェクト
    '' col_name: ソートをかける基準の列名
    '' order: 0; 昇順　1; 降順
    
    Dim col_num As Long
    Dim col_type As String
    Dim ar() As Variant
    Dim dic_list, dic_all_list, new_dic As Object
    Dim ltext As Variant
    Dim qs  As Variant
    Dim i As Long
    Dim ikey As Variant
    Dim wlist As Variant
    Dim keylist As Variant
    Dim anum As Long
    
    ''　列番号の検索
    col_num = WorksheetFunction.Match(col_name, dic.item("name"), 0) - 1
    '' データ型の検索
    col_type = dic.item("type")(col_num)
    '' 対象列データの取得
    
    Set dic_list = Set_Dictionary
    Set dic_all_list = Set_Dictionary
    For Each ikey In dic
        'If ikey <> "primary_key" Or ikey <> "name" Or ikey <> "type" Then
        Select Case ikey
            Case "primary_key", "name", "type"
            
            Case Else
                If dic_list.exists(dic.item(ikey)(col_num)) = False Then
                    dic_list.Add dic.item(ikey)(col_num), Text_Format(dic.item(ikey)(col_num), col_type)
                End If
                dic_all_list.Add ikey, Text_Format(dic.item(ikey)(col_num), col_type)
        End Select
        'End If
    Next

    ltext = dic_all_list.items
    
    '' 並び替えの必要がなければ処理を終了
    If dic_list.Count < 2 Then
        Debug.Print "並び替えを必要とするデータが存在しませんでした。"
        Set Sort_Data_Table = dic
        Exit Function
    End If
    
    '' quicksortで並び替え
    qs = dic_list.items
    qs = QuickSort(qs, LBound(qs), UBound(qs), order)
    
    ''　元のリストを並び変える行番号を検索
    wlist = Which_List(ltext, qs)
    
    '' 出力用のdictionary作成
    Set new_dic = Set_Dictionary
    
    '' ヘッダー部分を登録
    For Each ikey In dic
        If ikey = "primary_key" Or ikey = "type" Or ikey = "name" Then
            new_dic.Add ikey, dic.item(ikey)
        End If
    Next
    
    
    keylist = dic.keys
    For i = LBound(keylist) To UBound(keylist)
        If keylist(i) = "name" Then
            anum = i + 1
            Exit For
        End If
    Next
    
    For Each ikey In wlist
        new_dic.Add keylist(ikey + anum), dic.item(keylist(ikey + anum))
    Next
    
    Set Sort_Data_Table = new_dic
    
End Function
Private Function Which_List(ByVal target As Variant, ByVal data As Variant) As Variant
    '' targetで指定したリストの位置を番号で反す
    '' target, dataともに1次元配列 targetが文字列の場合は配列に変換して検索する
    
    Dim num, i As Long
    Dim dic As Object
    Dim ikey As Variant
    Dim dkey As Variant
    
    If Search_Dimentions(target) = 0 Then
        target = Array(target)
    End If
    
    Set dic = Set_Dictionary
    
    For Each ikey In data
        num = LBound(target)
        For Each dkey In target
            If ikey = dkey Then
                dic.Add CStr(num), num
            End If
            num = num + 1
        Next
    Next
    
    Which_List = dic.items
    
End Function

Private Function Format_List() As Variant
    Format_List = Array("日付", "時間", "数値", "文字列")
End Function
Private Function Text_Format(ByVal text As String, ByVal text_type As String) As Variant

    Dim output As Variant
    
    Select Case text_type
        Case "日付", "時間"
            output = CDate(text)
        Case "数値"
            output = CDbl(text)
        Case Else
            output = CStr(text)
    End Select
    
    Text_Format = output
    
End Function
Private Function Input_Format(ByVal text As String, ByVal text_type As String) As String
    On Error Resume Next
    Dim output As Variant
    
    Select Case text_type
        Case "日付"
            output = Format(text, "yyyy-MM-dd")
        Case "時間"
            output = Format(text, "HH:mm:SS")
        Case "数値"
            output = CDbl(text)
        Case Else
            output = CStr(text)
    End Select
    
    Input_Format = output
End Function
Private Function Get_Table_List() As Variant
    '' 既に作成されているテーブル名のリストを反す
    
    Dim buf As String
    Dim dtlist As Variant
    
    buf = dir(working_dir & "\*.txt")
    dtlist = ""
    
    Do While buf <> ""
        dtlist = dtlist & WorksheetFunction.Substitute(buf, ".txt", "", 1) & vbTab
        buf = dir()
    Loop
    
    If right(dtlist, 1) = vbTab Then
        dtlist = left(dtlist, Len(dtlist) - 1)
    End If
    
    Get_Table_List = Split(dtlist, vbTab)
    
End Function
Public Function select4dic(ByVal data_object As Object, ByVal select_column_names As Variant) As Object
    Set select4dic = Select_Columns(data_object, select_column_names)
End Function

Private Function Select_Columns(ByVal dic As Object, ByVal col_name As Variant) As Object
    '' namesに指定したデータの項目名(列名)だけを選択して出力する
    '' namesは1次元Array
    '' primary keyの列名はcol_nameに必ず入れること
    
    
    
    Dim dt, New_dt() As Variant
    Dim new_dic As Object
    Dim ch_dic As Object
    Dim i, j As Long
    Dim ar() As Variant
    Dim pk As Long
    Dim pk_name As String
    Dim check_names As String
    Dim ckey As Variant
    
    '' 列名の確認
    check_names = join(col_name, vbTab)
    If Not InStr(check_names, "NAME") > 0 Then
        col_name = Split("NAME" & vbTab & check_names, vbTab)
    End If
    
    '' primary key の確認
    pk = CLng(dic.item("primary_key")(1))
    pk_name = dic.item("name")(pk)
    pk = WorksheetFunction.Match(pk_name, col_name, 0) - 1
    
    '' 配列に変換
    dt = Dictionary_to_Array(dic)
    
    '' 配列を列名で列単位にまとめる
    Set new_dic = Set_Dictionary
    For j = LBound(dt, 2) To UBound(dt, 2)
        ReDim ar(LBound(dt, 1) To UBound(dt, 1))
        For i = LBound(dt, 1) To UBound(dt, 1)
            ar(i) = dt(i, j)
        Next
        '' dictionaryに登録
        new_dic.Add ar(2), ar
    Next
    
    '' 列をピックアップ
    Set ch_dic = Set_Dictionary
    For Each ckey In col_name
        If new_dic.exists(ckey) Then
            ch_dic.Add ckey, new_dic.item(ckey)
        End If
    Next
    
    '' dictionaryを配列に戻す
    ReDim New_dt(LBound(dt, 1) To UBound(dt, 1), 0 To ch_dic.Count - 1)
    j = 0
    For Each ckey In ch_dic
        ar = ch_dic.item(ckey)
        For i = LBound(ar) To UBound(ar)
            New_dt(i, j) = ar(i)
        Next
        j = j + 1
    Next
    
    ''配列をdictionaryに変換して元のデータを上書き
    Set new_dic = Array_to_Dictionary(New_dt, pk)
    Set Select_Columns = OverWrite_DataTable(dic, new_dic)
    
End Function
Private Function Shrink_Array(ByVal dt As Variant, ByVal remove_num As Long) As Variant
    '' 1次元配列から指定した番号の列を削除
    
    Dim ar() As Variant
    Dim num As Long
    Dim i As Long
    
    If Search_Dimentions(dt) <> 1 Then
        Debug.Print "入力データが1次元配列ではありません。"
        Exit Function
    End If
    
    ReDim ar(LBound(dt) To UBound(dt) - 1)
    
    For i = LBound(dt) To UBound(dt)
        If i < remove_num Then
            ar(i) = dt(i)
        ElseIf i > remove_num Then
            ar(i - 1) = dt(i)
        End If
    Next
    
    Shrink_Array = ar
    
End Function
Public Function cbind(ByVal dic As Object, ByVal join_key As String, ByVal join_dic As Object) As Object
    '' データを横に結合
    Set cbind = Left_Join(dic, join_key, join_dic)
End Function
Public Function rbind(ByVal dic As Object, ByVal union_dic As Object) As Object
    '' データを縦に結合
    Set rbind = Under_Union(dic, union_dic)
End Function
Private Function Under_Union(ByVal dic As Object, ByVal add_dic As Object) As Object
    '' データを縦に結合
    
        '' 既存のデータテーブルに新たな行を追加
    '' add_data: Dictionary型を内包する配列型(array)
    
    Dim new_dic As Object
    Dim cName As Variant
    Dim ncol As Long
    Dim pk As Long
    Dim ndim As Long
    Dim nrow As Long
    Dim dt, New_dt As Variant
    Dim add_name As Object
    Dim i As Long
    Dim key_name As Variant
    Dim d As Variant
    Dim ikey As Variant
    Dim dkey As Variant
    Dim d_name As Variant
    Dim d_d As Variant
    
    '' 列名取得
    cName = dic.item("name")
    
    '' プライマリーキー列番号
    pk = dic.item("primary_key")(1)
    
    '' 列名をもとにDictionaryを作成
    Set new_dic = Set_Dictionary
    For Each key_name In cName
        new_dic.Add key_name, ""
    Next
    
    '' 結合するデータから列名取得
    Set add_name = Set_Dictionary
    '' 列名を取得
    d_name = add_dic.item("name")
    For Each key_name In d_name
        add_name.Add key_name, ""
    Next

    
    '' 新規データを登録
    For Each d In add_dic
        '' ヘッダー部分と登録済みのキーは除外
        If d = "primary_key" Or d = "name" Or d = "type" Or dic.exists(d) Then GoTo Continue
        
        '' データを配列に取り込む
        d_d = add_dic.item(d)
        For i = LBound(d_d) To UBound(d_d)
            add_name.item(d_name(i)) = d_d(i)
        Next
        
        ''登録用dictionaryを修正
        For Each ikey In add_name
            If new_dic.exists(ikey) Then
                new_dic.item(ikey) = add_name.item(ikey)
            End If
        Next

        '' データ行を追加, 既存の場合は上書き
        dic.Add d, new_dic.items

        '' 既存のデータを一度消去
        For Each dkey In new_dic
            new_dic.item(dkey) = ""
        Next
        For Each dkey In add_name
            add_name.item(dkey) = ""
        Next
        d_d = ""
        
Continue:
    Next
    
    ''データを保存
    Set Under_Union = dic
    
End Function
Private Function Left_Join(ByVal dic As Object, ByVal key_name As String, _
                    ByVal add_dic As Object) As Object
    '' dicのkey_name列をキーとしてadd_dicの情報を右に結合する
    
    Dim num, anum As Long
    Dim add_table As Variant
    Dim New_table() As Variant
    Dim new_dic As Object
    Dim ar() As Variant
    Dim bname, aname As Variant
    Dim i, j As Long
    Dim dummy As String
    Dim ikey As Variant
    Dim d As Variant
    Dim skey As Variant
    
    bname = dic.item("name")
    '' 指定する列番号の検索
    num = WorksheetFunction.Match(key_name, bname, 0) - 1
    
    '' 追加するデータも同様に列を検索し、該当列は削除
    aname = add_dic.item("name")
    anum = WorksheetFunction.Match(key_name, aname, 0) - 1
    '' 配列に変換
    add_table = Dictionary_to_Array(add_dic)
    ''dictionary型に変換
    Set new_dic = Set_Dictionary
    ReDim ar(LBound(add_table, 2) To UBound(add_table, 2) - 1)
    'ReDim New_table(LBound(add_table, 1) To UBound(add_table, 1), LBound(add_table, 2) To UBound(add_table, 2) - 1)
    
    For i = LBound(add_table, 1) To UBound(add_table, 1)
        For j = LBound(add_table, 2) To UBound(add_table, 2)
            If j < anum Then
                ar(j) = add_table(i, j)
            ElseIf j > anum Then
                ar(j - 1) = add_table(i, j)
            End If
        Next
        '' データを登録
        Select Case ar(0)
            Case "TYPE"
                new_dic.Add "type", Shrink_Array(ar, 0)
            Case "NAME"
                new_dic.Add "name", Shrink_Array(ar, 0)
            Case Else
                new_dic.Add add_table(i, anum), Shrink_Array(ar, 0)
        End Select
        '' 格納したデータを消去
        For j = LBound(ar) To UBound(ar)
            ar(j) = ""
        Next
        
    Next
    
    dummy = String(UBound(aname) - 1, vbTab)
    
    '' キーをもとに結合
    For Each ikey In dic
        Select Case ikey
            Case "primary_key"
                GoTo Continue
                
            Case "type", "name"
                d = dic.item(ikey)
                skey = d(num)
                dic.item(ikey) = Split(join(d, vbTab) & vbTab & join(new_dic.item(ikey), vbTab), vbTab)
                
            Case Else
                d = dic.item(ikey)
                skey = d(num)
                If new_dic.exists(skey) Then
                    dic.item(ikey) = Split(join(d, vbTab) & vbTab & join(new_dic.item(skey), vbTab), vbTab)
                Else
                    dic.item(ikey) = Split(join(d, vbTab) & vbTab & dummy, vbTab)
                End If
        End Select
        
Continue:
    Next
    
    
    Set Left_Join = dic
    ''Left_Join = Dictionary_to_Array(dic)
    
End Function
Private Function Insert_New_Column_To_DataTable(ByVal DataTable_name As String, ByVal num As Long, _
                    ByVal col_type As Variant, ByVal col_name As Variant) As Object
    '' データテーブルに列を追加
    '' DataTable_name: DT名
    '' numで指定した列番号のあとに新規列を追加
    '' col_type: 追加する列のデータ型
    '' col_name: 追加する列名
    '' col_typeとcol_nameはデータ数が同じである必要があります。
    
    Dim pk As Long
    Dim New_dt() As Variant
    Dim ncol As Long
    Dim new_dic As Object
    Dim output As Object
    Dim dic As Object
    Dim n_check As Variant
    Dim ch As Variant
    Dim dt As Variant
    Dim i, j As Long
    
    If TypeName(col_type) = "String" And TypeName(col_name) = "String" Then
        col_type = Array(col_type)
        col_name = Array(col_name)
    End If
    
    If UBound(col_type) <> UBound(col_name) Then
        MsgBox "col_nameとcol_typeのデータ数が違います。" & vbCrLf & _
                "データ数を揃えてください。", vbInformation
        Exit Function
    End If
    
    ncol = UBound(col_type) + 1
    
    '' データを取得
    Set dic = Read_DataTable(DataTable_name)
    
    '' 列名の重複チェック
    n_check = join(dic.item("name"), vbTab)
    For Each ch In col_name
        If InStr(n_check, ch) > 0 Then
            Debug.Print "登録しようとしている列名に重複があります: " & ch
            Exit Function
        End If
    Next
    
    '' 配列化
    dt = Dictionary_to_Array(dic)
    
    
    '' プライマリーキー列
    pk = dic.item("primary_key")(1)
    If pk > num Then
        dic.item("primary_key")(1) = pk + UBound(col_type) + 1
    End If
    
    '' 列を追加
    ReDim New_dt(LBound(dt, 1) To UBound(dt, 1), LBound(dt, 2) To UBound(dt, 2) + UBound(col_type) + 1)
    
    For j = LBound(dt, 2) To UBound(dt, 2)
        '' 目的の列のあとに列を追加
        If j <= num Then
            For i = LBound(New_dt, 1) To UBound(New_dt, 1)
                New_dt(i, j) = dt(i, j)
            Next
        Else
            For i = LBound(New_dt, 1) To UBound(New_dt, 1)
                New_dt(i, j + ncol) = dt(i, j)
            Next
            
        End If
    Next
    
    '' 追加列情報を登録
    For j = LBound(col_type) To UBound(col_type)
        New_dt(LBound(New_dt, 1), num + 1 + j) = col_type(j)
        New_dt(LBound(New_dt, 1) + 1, num + 1 + j) = col_name(j)
    Next
    
    '' 配列からディクショナリーに変換
    Set new_dic = Array_to_Dictionary(New_dt, pk)
    
    '' データの上書き
    Set output = OverWrite_DataTable(dic, new_dic)
    
    Set Insert_New_Column_To_DataTable = output
    
    '' データの保存
    Call Write_DataTable(DataTable_name, output)
    
    
End Function
Private Function OverWrite_DataTable(ByVal dic As Object, ByVal new_dic As Object) As Object
    Dim nkey As Variant
    
    For Each nkey In new_dic
        If dic.exists(nkey) Then
            dic.item(nkey) = new_dic.item(nkey)
        End If
    Next
    
    Set OverWrite_DataTable = dic
    
End Function
Private Function Array_to_Dictionary(ByVal dt As Variant, ByVal num As Long) As Object
    '' num: プライマリーキー番号
    '' dt: データテーブル
    
    Dim dic As Object
    Dim ar() As Variant
    Dim i, j As Long
    Set dic = Set_Dictionary
    
    ReDim ar(LBound(dt, 2) To UBound(dt, 2))
    
    For i = LBound(dt, 1) To UBound(dt, 1)
        '' 登録データを1次元に変換
        For j = LBound(dt, 2) To UBound(dt, 2)
            ar(j) = dt(i, j)
        Next
        '' データを登録
        Select Case ar(0)
            Case "TYPE"
                dic.Add "type", ar
            Case "NAME"
                dic.Add "name", ar
            Case Else
                dic.Add CStr(ar(num)), ar
        End Select
        
        '' 格納したデータを消去
        For j = LBound(dt, 2) To UBound(dt, 2)
            ar(j) = ""
        Next
        
    Next
    
    Set Array_to_Dictionary = dic
    
End Function

'Private Function Update_DataTable(ByVal DataTable_name As String, _
'                                    ByVal add_data As Variant) As Object
'    '' 既存のデータテーブルに新たな行を追加
'    '' add_data: Dictionary型を内包する配列型(array)
'
'    Dim dic, New_dic As Object
'    Dim cname As Variant
'    Dim ncol As Long
'    Dim pk As Long
'    Dim ndim As Long
'    Dim nrow As Long
'    Dim dt, New_dt As Variant
'    Dim ctype As Variant
'    Dim dic_type As Object
'    Dim set_date As String
'    Dim i As Long
'    Dim key_name As Variant
'    Dim d As Variant
'    Dim ikey As Variant
'    Dim push As Variant
'    Dim dkey As Variant
'
'
'    '' 登録日作成
'    set_date = Format(Now(), "yyyy-mm-dd hh:MM:ss")
'    '' 一覧データ取得
'    Set dic = Read_DataTable(DataTable_name)
'    '' 列名取得
'    cname = dic.item("name")
'    '' データタイプ取得
'    ctype = dic.item("type")
'    Set dic_type = Set_Dictionary
'    For i = LBound(cname) To UBound(cname)
'        dic_type.Add cname(i), ctype(i)
'    Next
'
'    '' プライマリーキー列番号
'    pk = dic.item("primary_key")(1)
'
'    '' 列名をもとにDictionaryを作成
'    Set New_dic = Set_Dictionary
'    For Each key_name In cname
'        New_dic.Add key_name, ""
'    Next
'
'    ''データ型を変換
'    '' すでに変換されている場合は除外、2次元配列であれば変換
'    If Search_Dimentions(add_data) = 2 Then
'        add_data = Get_Table_To_Array(add_data)
'    End If
'
'    '' 新規データを登録
'    For Each d In add_data
'        ''データを修正
'        For Each ikey In d
'            If New_dic.exists(ikey) Then
'                New_dic.item(ikey) = Input_Format(d.item(ikey), dic_type.item(ikey))
'            End If
'        Next
'        New_dic.item("NAME") = set_date
'        '' データ行を追加, 既存の場合は上書き
'        push = New_dic.items
'        If dic.exists(push(pk)) Then
'            dic.item(push(pk)) = push
'        Else
'            dic.Add push(pk), push
'        End If
'
'        '' 既存のデータを一度消去
'        For Each dkey In New_dic
'            New_dic.item(dkey) = ""
'        Next
'        push = ""
'    Next
'
'    ''データを保存
'    ''Call Write_DataTable(DataTable_name, New_dic)
'    Set Update_DataTable = dic
'
'End Function
Private Function Update_DataTable(ByVal DataTable_name As String, _
                                    ByVal add_data As Variant) As Object
    '' 既存のデータテーブルに新たな行を追加
    '' add_data: Dictionary型を内包する配列型(array)
    
    Dim dic, new_dic As Object
    Dim cName As Variant
    Dim ncol As Long
    Dim pk As Long
    Dim ndim As Long
    Dim nrow As Long
    Dim dt, New_dt As Variant
    Dim ctype As Variant
    Dim dic_type As Object
    Dim set_date As String
    Dim i As Long
    Dim key_name As Variant
    Dim d As Variant
    Dim ikey As Variant
    Dim push As Variant
    Dim dkey As Variant
    Dim pkv As String
    Dim set_values As Variant
    
    '' 登録日作成
    set_date = Format(Now(), "yyyy-mm-dd hh:MM:ss")
    '' 一覧データ取得
    Set dic = Read_DataTable(DataTable_name)
    '' 列名取得
    cName = dic.item("name")
    '' データタイプ取得
    ctype = dic.item("type")
    Set dic_type = Set_Dictionary
    For i = LBound(cName) To UBound(cName)
        dic_type.Add cName(i), ctype(i)
    Next
    
    '' プライマリーキー列番号
    pk = dic.item("primary_key")(1)
    
    '' 列名をもとにDictionaryを作成
'    Set new_dic = Set_Dictionary
'    For Each key_name In cName
'        new_dic.Add key_name, ""
'    Next
    
    ''データ型を変換
    '' すでに変換されている場合は除外、2次元配列であれば変換
    If Search_Dimentions(add_data) = 2 Then
        add_data = Get_Table_To_Array(add_data)
    End If
    
    '' 新規データを登録
    For Each d In add_data
        ''キーを取得
        pkv = d.item(cName(pk))
        '' 既存のデータ検索
        If dic.exists(pkv) Then
            set_values = dic.item(pkv)
        
            '' 初期値として既存のデータを指定
            Set new_dic = Set_Dictionary
            i = LBound(set_values)
            For Each key_name In cName
                new_dic.Add key_name, set_values(i)
                i = i + 1
            Next
        Else
            '' 初期値として既存のデータを指定
            Set new_dic = Set_Dictionary
            For Each key_name In cName
                new_dic.Add key_name, ""
            Next
        End If
        
        ''データを修正
        For Each ikey In d
            
            If new_dic.exists(ikey) Then
                new_dic.item(ikey) = Input_Format(d.item(ikey), dic_type.item(ikey))
            End If

        Next
        new_dic.item("NAME") = set_date
        
        '' データ行を追加, 既存の場合は上書き
        push = new_dic.items
        If dic.exists(pkv) Then
            dic.item(pkv) = push
        Else
            dic.Add pkv, push
        End If
'
'        '' 既存のデータを一度消去
'        For Each dkey In new_dic
'            new_dic.item(dkey) = ""
'        Next
        push = ""
    Next
    
    ''データを保存
    ''Call Write_DataTable(DataTable_name, New_dic)
    Set Update_DataTable = dic
    
End Function
Private Function Dictionary_to_Array(ByVal dic As Object) As Variant
    '' Dictionary → 二次元配列に変換
    '' dicはDictionary型データを指定
    
    Dim data() As Variant
    Dim nrow, ncol As Long
    Dim key_names As Variant
    Dim i, j As Long
    Dim ld As Variant
    
    nrow = dic.Count - 1
    
'    If nrow < 2 Then
'        Debug.Print "データテーブルにデータが登録されていません。"
'        Exit Function
'    End If
    
    key_names = dic.keys
    
    ncol = UBound(dic.item(key_names(2)))
    
    ReDim data(1 To nrow, ncol)
    
    For i = 1 To nrow
    
        ld = dic.item(key_names(i))
        
        For j = 0 To ncol
            data(i, j) = ld(j)
        Next
        
        ld = ""
    Next
    
    
    Dictionary_to_Array = data


End Function
Private Function Read_DataTable(ByVal DataTable_name As String) As Object
    '' データテーブルを取り込み、Dictionary型で出力する関数
    '' 区切り文字はタブ、改行文字はVbCr
    
    Dim n As Integer
    Dim i As LongPtr
    Dim text As String
    Dim dic As Object
    Dim pk As Long
    Dim data As Variant
    
    
    Set dic = Set_Dictionary
    n = FreeFile
    
    Open DT_name(DataTable_name) For Input As #n
        i = 1
        Do While Not EOF(n)
            Line Input #n, text
            
            If text = "" Then GoTo Continue
            
            Select Case i
                Case 1
                    dic.Add "primary_key", Split(text, vbTab)
                    pk = CLng(dic.item("primary_key")(1))
                Case 2
                    dic.Add "type", Split(text, vbTab)
                Case 3
                    dic.Add "name", Split(text, vbTab)
                Case Else
                    data = Split(text, vbTab)
                    dic.Add data(pk), data
            End Select
            
Continue:
            i = i + 1
        Loop
    
    Close #n
    
    
    Set Read_DataTable = dic
    
End Function
Private Sub Add_Data_Table(ByVal name As String, ByVal KEY As LongPtr, _
                        ByVal col_type As Variant, ByVal col_name As Variant)
    '' データテーブルを作成
    '' name: DTの名称
    '' KEY: Primary Keyの列番号 1～
    '' col_type: 追加する列のデータ型
    '' col_name: 追加する列名
    '' col_typeとcol_nameはデータ数が同じである必要があります。
    
    Dim n As Integer
    Dim text As String
    Dim filename As String
    Dim msg As Variant
    
    If TypeName(col_type) = "String" And TypeName(col_name) = "String" Then
        col_type = Array(col_type)
        col_name = Array(col_name)
    End If
    
    If UBound(col_type) <> UBound(col_name) Then
        MsgBox "col_nameとcol_typeのデータ数が違います。" & vbCrLf & _
                "データ数を揃えてください。", vbInformation
        Exit Sub
    End If
    
    n = FreeFile
    text = "PRIMARY KEY: " & vbTab & KEY & vbCr & _
            "TYPE" & vbTab & join(col_type, vbTab) & vbCr & _
            "NAME" & vbTab & join(col_name, vbTab) & vbCr
    filename = DT_name(name)
    If dir(filename, vbDirectory) <> "" Then
        msg = MsgBox(name & ": このテーブルは既に作成されています。" & vbCrLf & _
                        "続行するとテーブルが削除されます。", vbYesNo)
        If msg = 7 Then Exit Sub
    End If
    Open filename For Output As #n
        Print #n, text
    Close #n
    
    Debug.Print "Create New Data Table: " & filename
    
End Sub
Public Sub WriteTableWithoutHeader(ByVal data_table_name As String, ByVal data_object As Variant)
    Call Write_DataTable(data_table_name, data_object)
    Call Make_Backup_File(data_table_name)
End Sub
Public Function ReadTableWithoutHeader(ByVal data_table_name As String, _
                                        Optional read_type As readtype = readtype.RowHead, _
                                        Optional separate_key As SepKey = SepKey.keyTab) As Object
    '' データテーブルを取り込み、Dictionary型で出力する関数
    '' 区切り文字はタブ、改行文字はVbCr
    
    Dim n As Integer
    Dim i As LongPtr
    Dim text As String
    Dim dic As Object
    Dim pk As Long
    Dim data As Variant
    
    
    Set dic = Set_Dictionary
    n = FreeFile
    If InStr(data_table_name, "\") = 0 Then data_table_name = DT_name(data_table_name)
        
    Open data_table_name For Input As #n
        i = 1
        Do While Not EOF(n)
            Line Input #n, text
            
            If text = "" Then GoTo Continue
            
            If separate_key = keyTab Then
                data = Split(text, vbTab)
            ElseIf separate_key = keyComma Then
                data = Split(text, ",")
            End If
            
            If read_type = readtype.RowHead Then
                dic.Add CStr(data(0)), data
            ElseIf read_type = readtype.Number Then
                dic.Add CStr(i), data
            End If
Continue:
            i = i + 1
        Loop
    
    Close #n
    
    
    Set ReadTableWithoutHeader = dic
    
End Function
Private Sub Write_DataTable(ByVal DataTable_name As String, ByVal Write_data As Variant)
    '' データテーブルを上書き保存
    '' DataTable_name: テーブル名
    '' Write_data: 上書きするDictionary型データ
    On Error GoTo Error
    
    Dim n As Integer
    Dim text As String
    Dim filename As String
    Dim dic As Object
    Dim cnt As Long
    Dim ikey As Variant
    
    cnt = 1
    
    Select Case VarType(Write_data)
        Case 9 ''Object
            Set dic = Set_Dictionary
            For Each ikey In Write_data
                dic.item(ikey) = join(Write_data.item(ikey), vbTab)
            Next
            text = join(dic.items, vbCr)
        
        Case 8 '' String
            text = Write_data
        
        Case 8192 ''Array
            
    End Select
    filename = DT_name(DataTable_name)

Write_Run:
    n = FreeFile
    Open filename For Output Lock Write As #n
        Print #n, text
    Close #n
    GoTo Continue
    
Error:
    ''openされていれば閉じる処理を追加
    If Err.Number > 0 Then Close #n
    
    If cnt > 5 Then
        MsgBox "他のユーザーがファイルを開いています。少し待ってから再開してください。" & vbCrLf & _
                "※　[OK]ボタンを押すと処理が再開します。", vbCritical
    Else
        Application.Wait Now() + TimeValue("00:00:01")
    End If
    cnt = cnt + 1
    GoTo Write_Run
    
Continue:
    Debug.Print "Save Data Table: " & filename
    
End Sub
Public Sub MakeDirectory(ByVal dir_name As String)
    Call Create_New_Folder(dir_name)
End Sub
Private Sub Create_New_Folder(ByVal fold_name As String)
    Dim fol_name As String
    Dim check_dir As String
    
    If fold_name = "" Then
        fol_name = working_dir
    Else
        fol_name = Join_Path(working_dir, fold_name)
    End If
    
    check_dir = dir(fol_name, vbDirectory)
    If check_dir = "" Then
        mkdir fol_name
        Debug.Print "Make directory: " & fol_name
    End If
    
End Sub
Private Sub Set_New_Data_Base()
    '' 新規でデータベースフォルダを作成
    
    Dim dir_name As Variant
    
'' サブディレクトリの作成 ==============================
    Call Create_New_Folder("")
    
'    For Each dir_name In Array("Backups", "Report", "Picture")
'        Call Create_New_Folder(dir_name)
'    Next
    
    Call Create_New_Folder("Backups")

End Sub
Private Function Search_Dimentions(ByVal dt As Variant) As Long
    On Error GoTo Continue
    Dim i As Integer
    Dim nmax, nmin As Variant
    Dim dims As String
    
    i = 1
    nmax = Empty
    
    Do
        nmax = UBound(dt, i)
        If IsEmpty(nmax) Then GoTo Continue
        i = i + 1
    Loop
    
Continue:
    
    Search_Dimentions = i - 1
    
End Function
Public Function Join_Path(ByVal path As String, ByVal name As String) As String
    Join_Path = path & "\" & name
End Function
Private Function DT_name(ByVal name As String)
    DT_name = Join_Path(working_dir, name & ".txt")
End Function
Public Function Set_Dictionary() As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Set Set_Dictionary = dic
End Function
Private Function QuickSort(ByVal data As Variant, ByVal L As Long, _
                    ByVal U As Long, ByVal order As Integer) As Variant
    '' 逆順番に並び替え
    '' data: 1次元配列
    '' order: 0:昇順　1:降順
    
    Dim i As Long
    Dim j As Long
    Dim S As Variant
    Dim Tmp As Variant
    
Continue:

    S = data(Int((L + U) / 2))
    i = L
    j = U
    Do
        If order = 0 Then
            Do While data(i) < S
                i = i + 1
            Loop
            
            Do While data(j) > S
                j = j - 1
            Loop
        ElseIf order = 1 Then
            Do While data(i) > S
                i = i + 1
            Loop
            
            Do While data(j) < S
                j = j - 1
            Loop
        End If
        If i >= j Then Exit Do
        
        Tmp = data(i)
        data(i) = data(j)
        data(j) = Tmp
        i = i + 1
        j = j - 1
    Loop
    If L < i - 1 Then
        U = i - 1
        L = L
        GoTo Continue
    End If
    If U > j + 1 Then
        L = j + 1
        U = U
        GoTo Continue
    End If
      
    QuickSort = data
    
End Function

'Private Function wd() As String
'    wd = working_dir
'End Function

Property Get WorkDir() As String
    WorkDir = working_dir
End Property
Property Let WorkDir(ByVal working_directory As String)
    working_dir = working_directory & "\data"
End Property

Private Sub Class_Initialize()
    working_dir = ThisWorkbook.Worksheets(1).Cells(1, 1).Value & "\data"
End Sub
