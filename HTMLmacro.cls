VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTMLmacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Sub sample(ByVal data_table_name As String)
    '' �f�[�^��I������HTML�t�H�[�}�b�g�ŕ\��
    
    Dim num As Long
    Dim dic As Object
    Dim nrow As Variant
    Dim data As Object
    Dim ikey As Variant
    Dim body As String
    Dim name, items As Variant
    Dim df As Object
    Dim TSQL As New textSQL
    Dim HTML As New HTMLmacro
    
    '' �f�[�^�e�[�u���捞��
    Set data = TSQL.SELECT_Table(data_table_name)
    
    '' �Y������f�[�^�����o��
    body = ""
    name = data.item("name")
    For Each ikey In dic.keys
        Set df = TSQL.Set_Dictionary
        items = data.item(ikey)
        '' �����ɍ��v������̂�o�^(�S�ĕ\�ŏo�͂���ꍇ)
        For i = LBound(name) + 1 To UBound(name)
            df.Add name(i), Array(items(i))
        Next
        
        '' HTML�̃t�H�[�}�b�g�w��
        With HTML
            body = body & _
                .article( _
                    .h1(.font(ikey, "red")) & _
                    .table(df, "wide"))
            
        End With
    Next
    
    '' HTML�o��
    text = HTML.HTML("�e�X�g - HTML�o��", body)
    
    Call HTML.PrintHTMLandOpenByBrowser(text)
    
End Sub
Public Sub PrintHTMLandOpenByBrowser(ByVal text As String)
    '' HTML�t�@�C���Ƃ��ďo��&�\��
    
    Dim filepath, filename As String
    Dim fso As Object
    Dim TSQL As New textSQL
    
    '' �t�@�C���ۑ��ꏊ
    
    filename = TSQL.WorkDir & "\" & Format(Now(), "yyyyMMdd_HHmmSS") & ".html"
    
    '' html�t�@�C�����쐬
    Set fso = CreateObject("ADODB.Stream")
    With fso
        '' �e�L�X�g�f�[�^�̎捞��
        .mode = 3 ''
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        .WriteText text, 1
        '' �ۑ����ĕ���
        
        .SaveToFile filename, adSaveCreateOverWrite
        .Close
    End With
    
    '' Edge���N�����ăt�@�C�����J��
    MsgBox "Microsoft Edge�Ŋm�F�p��ʂ�\�����܂��B" & vbCrLf & _
            "�������ꍇ�́A�u���E�U�̈���@�\���g���Ă��������B", vbInformation
    CreateObject("WScript.Shell").RUN "msedge.exe -url " & filename
    
    '' �쐬�����t�@�C�����폜
    Application.Wait Now() + TimeValue("00:00:05")
    MsgBox "Excel��ʂɖ߂�܂��B"
    Kill filename
    
End Sub
Function bgc(ByVal text As String)
    bgc = WorksheetFunction.Substitute(text, ">", " class=table-bgc>", 1)
End Function
Public Function font(ByVal text As String, ByVal color As String) As String
    font = "<font color=" & color & ">" & text & "</font>"
End Function
Public Function span(ByVal text As String, ByVal rowspan As Integer, ByVal colspan As Integer) As String
    '' �\�g�݂ŃZ���̌������s��
    Dim output As String
    output = WorksheetFunction.Substitute(text, "<td>", "<td rowspan=" & rowspan & " colspan=" & colspan & ">", 1)
    output = WorksheetFunction.Substitute(output, "<th>", "<th rowspan=" & rowspan & " colspan=" & colspan & ">", 1)
    span = output
    
End Function
Public Function th(ByVal text As String) As String
    th = "<th>" & text & "</th>" & vbLf
End Function
Public Function tr(ByVal text As String) As String
    tr = "<tr>" & text & "</tr>" & vbLf
End Function
Public Function td(ByVal text As String) As String
    td = "<td>" & text & "</td>" & vbLf
End Function
Public Function table(ByVal df As Variant, ByVal face As String) As String
    '' df ��Dictionary�^
    '' Keys = ��
    '' items = �v�f
    '' items�̃f�[�^���͍��킹�Ă��������B
    '' face: wide = key, items ...
    ''              key, items ...
    ''              key, items ...
    '' face: long = key, key, key...
    ''              items, items, items ...
    ''              items, items, items ...
    
    Dim text, linetxt  As String
    Dim KEY As Variant
    Dim n, i As Long
    
    Select Case face
        Case "wide"
            n = 0
            text = ""
            For Each KEY In df.keys
                If n = 0 Then
                    text = text & tr(th(KEY) & td(join(df(KEY), "</td><td>")))
                Else
                    text = text & tr(th(KEY) & td(join(df(KEY), "</td><td>")))
                End If
                
                n = n + 1
            Next
        Case "long"
            '' ���o��
            text = tr(th(join(df.keys, "</th><th>")))
            n = UBound(df(Split(join(df.keys, ","), ",")(0)))
            i = 0
            Do While i <= n
                linetxt = ""
                For Each KEY In df
                    linetxt = linetxt & td(df(KEY)(i))
                Next
                text = text & tr(linetxt)
                i = i + 1
            Loop
            
        
    End Select
    
    table = "<table>" & vbLf & _
            text & vbLf & _
            "</table>" & vbLf
End Function
Public Function ul(ByVal text As String) As String
    ul = "<ul>" & vbLf & _
            text & vbLf & _
            "</ul>" & vbLf
End Function
Public Function li(ByVal text As String) As String
    li = "<li>" & text & "</li>" & vbLf
End Function
Public Function list(ByVal data As Variant) As String
    '' data: array
    Dim text As String
    Dim tx As Variant
    
    text = ""
    For Each tx In data
        text = text & li(tx)
    Next
    
    list = ul(text)
    
End Function
Public Function right(ByVal text As String) As String
    right = WorksheetFunction.Substitute(text, ">", " style=""text-align:right"">", 1)
End Function
Public Function center(ByVal text As String) As String
    center = WorksheetFunction.Substitute(text, ">", " style=""text-align:center"">", 1)
End Function
Public Function left(ByVal text As String) As String
    left = WorksheetFunction.Substitute(text, ">", " style=""text-align:left"">", 1)
End Function
Public Function p(ByVal text As String) As String
    p = "<p>" & vbLf & _
            text & vbLf & _
            "</p>" & vbLf
End Function
Public Function h3(ByVal text As String) As String
    h3 = "<h3>" & vbLf & _
            text & vbLf & _
            "</h3>" & vbLf
End Function
Public Function h2(ByVal text As String) As String
    h2 = "<h2>" & vbLf & _
            text & vbLf & _
            "</h2>" & vbLf
End Function
Public Function h1(ByVal text As String) As String
    h1 = "<h1>" & vbLf & _
            text & vbLf & _
            "</h1>" & vbLf
End Function
Public Function article(ByVal text As String) As String
    article = "<article>" & vbLf & _
                text & vbLf & _
                "</article>" & vbLf & _
                "<div class=""pagebreak""></div>" & vbLf
End Function
Public Function HTML(ByVal title As String, ByVal body As String) As String
    Dim str As String
    
    
    str = _
    "<!DOCTYPE HTML>" & vbLf & _
    "<html lang = ""ja"">" & vbLf & _
    "<head>" & vbLf & _
    "<title>" & title & "</title>" & vbLf & _
    "<meta charset=""UTF-8"">" & vbLf & _
    "<style type=""text/css"">" & vbLf & _
    "table {border-collapse: collapse; width: 100%;}" & vbLf & _
    "th,td {border: solid 1px; padding: 10px;}" & vbLf & _
    "th {width: 25%;}" & vbLf & _
    ".pagebreak {break-after: page;}" & vbLf & _
    ".table-bgc {background: #FF3300;}" & _
    "</style>" & vbLf & _
    "</head>" & vbLf & _
    "<body>" & vbLf & _
    body & vbLf & _
    "</body>" & vbLf & _
    "</html>"
    
    
    HTML = str
    
End Function



