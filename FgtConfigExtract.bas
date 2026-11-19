Option Explicit
'==============================================================
' FortiGate コンフィグ抽出マクロ（汎用パーサー版）
'
' 使い方
'   1. VBE(Alt+F11) → 挿入 → 標準モジュール にこのコードを丸ごと貼り付け
'   2. RunFgtExtract を実行 → コンフィグファイルを選択
'   3. 「設定」シートA列のセクションごとにシートが作られる
'
' 「設定」シート（初回実行時に自動作成）
'   A列: 抽出するセクション名（例: firewall policy）※2行目以降
'   B列: 抽出件数（実行時に自動で書き込み）
'   D列: 出力しないキー（例: uuid, passwd）※2行目以降
'   F列: 最後に読み込んだファイルと日時
'
' 仕様
'   - config / edit / set / next / end の構造を汎用的に読むので、
'     セクションごとの専用ロジックは無し。列は登場した set のキーから自動生成
'   - 複数値（set srcaddr "A" "B"）はセル内改行で連結（引用符なしの値は空白区切り）
'   - VDOM構成（config vdom / config global）にも対応。VDOM列を自動追加
'   - edit を持たないセクション（system global 等）は1行で出力
'   - ネストした config は「親config名.キー」の列にまとめる
'     （ネスト内に edit がある場合は「edit名: 値」の形でセル内改行連結）
'==============================================================

Private Const SETTING_SHEET As String = "設定"
Private Const MAX_CELL_LEN As Long = 32000
Private Const MAX_COL_WIDTH As Double = 50
Private Const MAX_BUF_LINES As Long = 1000   ' 引用符が閉じない場合の安全弁

'==============================================================
' メイン
'==============================================================
Public Sub RunFgtExtract()
    Dim wsSet As Worksheet
    Dim targets As Object, excludes As Object
    Dim f As Variant, txt As String
    Dim result As Object, sec As Object
    Dim k As Variant, cnt As Long, shName As String

    On Error GoTo EH

    Set wsSet = GetSettingSheet()
    Set targets = ReadList(wsSet, 1)
    Set excludes = ReadList(wsSet, 4)
    If targets.Count = 0 Then
        MsgBox "「設定」シートのA列（2行目以降）に抽出するセクション名を入力してください。", vbExclamation
        Exit Sub
    End If

    f = Application.GetOpenFilename( _
        "コンフィグ (*.conf;*.txt;*.log;*.cfg),*.conf;*.txt;*.log;*.cfg,すべてのファイル (*.*),*.*", _
        , "FortiGateのコンフィグを選択")
    If VarType(f) = vbBoolean Then Exit Sub

    txt = ReadTextUtf8(CStr(f))
    Set result = ParseFgtConfig(txt, targets, excludes)

    Application.ScreenUpdating = False

    For Each k In targets.Keys
        cnt = 0
        shName = SafeSheetName(CStr(k))
        If result.Exists(k) Then
            Set sec = result(k)
            WriteSection shName, sec
            cnt = sec("rows").Count
        ElseIf SheetExists(shName) Then
            ' 前回の案件のデータが残らないようにクリア
            With ThisWorkbook.Worksheets(shName)
                If .AutoFilterMode Then .AutoFilterMode = False
                .Cells.Clear
                .Range("A1").Value = "(このコンフィグには該当セクションなし)"
            End With
        End If
        wsSet.Cells(targets(k), 2).Value = cnt
    Next k

    wsSet.Range("F2").Value = CStr(f)
    wsSet.Range("F3").Value = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    wsSet.Activate

    Application.ScreenUpdating = True
    MsgBox "抽出が完了しました。件数は「設定」シートのB列を確認してください。", vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Number & vbLf & Err.Description, vbCritical
End Sub

'==============================================================
' パーサー本体
'   戻り値: Dictionary( セクション名 → セクション(Dictionary) )
'   セクション: "rows"=Collection(レコード), "cols"=Dictionary(キー→列番号), "hasVdom"=Boolean
'   レコード:   Dictionary( "__vdom", "__edit", 各キー → 値 )
'==============================================================
Private Function ParseFgtConfig(ByVal txt As String, ByVal targets As Object, ByVal excludes As Object) As Object
    Dim result As Object
    Dim lines() As String
    Dim stKind() As String, stName() As String
    Dim depth As Long          ' 現在のネストの深さ
    Dim tDepth As Long         ' 抽出中セクションの config の深さ（0 = 抽出中でない）
    Dim tName As String
    Dim sec As Object, rec As Object
    Dim i As Long, j As Long, p As Long
    Dim ln As String, buf As String, bufLines As Long
    Dim cmd As String, rest As String
    Dim key As String, leaf As String, val As String
    Dim startIdx As Long, prefix As String, label As String

    Set result = CreateObject("Scripting.Dictionary")

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    ReDim stKind(0 To 100)
    ReDim stName(0 To 100)

    For i = 0 To UBound(lines)
        ' --- TeraTermログ等のゴミを除去 ---
        ln = Replace(lines(i), "--More--", "")
        ln = Replace(ln, Chr$(8), "")

        ' --- "..." が複数行にまたがる値（コメント・証明書等）を1行にまとめる ---
        If Len(buf) > 0 Then
            buf = buf & vbLf & ln
            bufLines = bufLines + 1
            If IsQuoteOpen(buf) And bufLines < MAX_BUF_LINES Then GoTo NextLine
            ln = buf
            buf = ""
            bufLines = 0
        ElseIf IsQuoteOpen(ln) Then
            buf = ln
            bufLines = 1
            GoTo NextLine
        End If

        ln = Trim$(Replace(ln, vbTab, " "))
        If Len(ln) = 0 Then GoTo NextLine
        If Left$(ln, 1) = "#" Then GoTo NextLine

        cmd = LCase$(FirstWord(ln))
        rest = Trim$(Mid$(ln, Len(cmd) + 1))

        Select Case cmd

        Case "config"
            depth = PushFrame(stKind, stName, depth, "config", NormalizeSpaces(LCase$(rest)))
            If tDepth = 0 Then
                If IsTopLevel(stKind, stName, depth - 1) Then
                    If targets.Exists(stName(depth)) Then
                        tDepth = depth
                        tName = stName(depth)
                        If Not result.Exists(tName) Then result.Add tName, NewSection()
                        Set sec = result(tName)
                        Set rec = Nothing
                    End If
                End If
            End If

        Case "edit"
            depth = PushFrame(stKind, stName, depth, "edit", JoinTokens(Tokenize(rest), " "))
            If tDepth > 0 Then
                If depth = tDepth + 1 Then
                    Set rec = NewRecord(sec, CurrentVdom(stKind, stName, depth), stName(depth))
                End If
            End If

        Case "next"
            If depth > 0 Then
                If stKind(depth) = "edit" Then
                    If tDepth > 0 Then
                        If depth = tDepth + 1 Then Set rec = Nothing
                    End If
                    depth = depth - 1
                End If
            End If

        Case "end"
            ' next が抜けていても壊れないよう、残っている edit を畳む
            Do While depth > 0
                If stKind(depth) <> "edit" Then Exit Do
                depth = depth - 1
            Loop
            If depth > 0 Then
                If depth = tDepth Then
                    tDepth = 0
                    tName = ""
                    Set sec = Nothing
                    Set rec = Nothing
                End If
                depth = depth - 1
            End If

        Case "set"
            If tDepth > 0 Then
                If depth >= tDepth Then
                    p = InStr(rest, " ")
                    If p = 0 Then
                        leaf = LCase$(rest)
                        val = ""
                    Else
                        leaf = LCase$(Left$(rest, p - 1))
                        ' "A" "B" 形式（オブジェクト名の列挙）はセル内改行、
                        ' 引用符なし（IP マスク / ping https 等）は空白区切りのまま
                        If InStr(rest, """") > 0 Then
                            val = JoinTokens(Tokenize(Mid$(rest, p + 1)), vbLf)
                        Else
                            val = JoinTokens(Tokenize(Mid$(rest, p + 1)), " ")
                        End If
                    End If

                    If Not excludes.Exists(leaf) Then
                        ' レコード直下（edit 直下 or edit を持たないセクション直下）より深い部分を判定
                        startIdx = tDepth + 1
                        If depth > tDepth Then
                            If stKind(tDepth + 1) = "edit" Then startIdx = tDepth + 2
                        End If

                        If rec Is Nothing Then
                            ' edit を持たないセクション（system global 等）
                            Set rec = NewRecord(sec, CurrentVdom(stKind, stName, tDepth), "-")
                        End If

                        prefix = ""
                        label = ""
                        For j = startIdx To depth
                            If stKind(j) = "config" Then
                                If Len(prefix) > 0 Then prefix = prefix & "."
                                prefix = prefix & stName(j)
                            Else
                                If Len(label) > 0 Then label = label & "/"
                                label = label & stName(j)
                            End If
                        Next j

                        key = leaf
                        If Len(prefix) > 0 Then key = prefix & "." & leaf
                        If Len(label) > 0 Then val = label & ": " & Replace(val, vbLf, " ")

                        AddField sec, rec, key, val
                    End If
                End If
            End If

        Case Else
            ' unset / append / プロンプト行などは無視
        End Select
NextLine:
    Next i

    Set ParseFgtConfig = result
End Function

'--------------------------------------------------------------
' 構造スタック
'--------------------------------------------------------------
Private Function PushFrame(stKind() As String, stName() As String, ByVal depth As Long, _
                           ByVal kind As String, ByVal nm As String) As Long
    depth = depth + 1
    If depth >= UBound(stKind) Then
        ReDim Preserve stKind(0 To depth * 2)
        ReDim Preserve stName(0 To depth * 2)
    End If
    stKind(depth) = kind
    stName(depth) = nm
    PushFrame = depth
End Function

' 抽出対象として扱える「トップレベル」か
'   素の構成: 深さ0 / VDOM構成: config vdom > edit X の直下 / config global の直下
Private Function IsTopLevel(stKind() As String, stName() As String, ByVal d As Long) As Boolean
    Select Case d
    Case 0
        IsTopLevel = True
    Case 1
        IsTopLevel = (stKind(1) = "config" And stName(1) = "global")
    Case 2
        IsTopLevel = (stKind(1) = "config" And stName(1) = "vdom" And stKind(2) = "edit")
    Case Else
        IsTopLevel = False
    End Select
End Function

Private Function CurrentVdom(stKind() As String, stName() As String, ByVal d As Long) As String
    CurrentVdom = ""
    If d >= 2 Then
        If stKind(1) = "config" And stName(1) = "vdom" And stKind(2) = "edit" Then
            CurrentVdom = stName(2)
            Exit Function
        End If
    End If
    If d >= 1 Then
        If stKind(1) = "config" And stName(1) = "global" Then CurrentVdom = "global"
    End If
End Function

'--------------------------------------------------------------
' データ構造
'--------------------------------------------------------------
Private Function NewSection() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "rows", New Collection
    d.Add "cols", CreateObject("Scripting.Dictionary")
    d.Add "hasVdom", False
    Set NewSection = d
End Function

Private Function NewRecord(ByVal sec As Object, ByVal vdom As String, ByVal nm As String) As Object
    Dim r As Object, rows As Collection
    Set r = CreateObject("Scripting.Dictionary")
    r.Add "__vdom", vdom
    r.Add "__edit", nm
    Set rows = sec("rows")
    rows.Add r
    If Len(vdom) > 0 Then sec("hasVdom") = True
    Set NewRecord = r
End Function

Private Sub AddField(ByVal sec As Object, ByVal rec As Object, ByVal key As String, ByVal val As String)
    Dim cols As Object
    Set cols = sec("cols")
    If Not cols.Exists(key) Then cols.Add key, cols.Count + 1
    If rec.Exists(key) Then
        rec(key) = rec(key) & vbLf & val
    Else
        rec.Add key, val
    End If
End Sub

'--------------------------------------------------------------
' 出力
'--------------------------------------------------------------
Private Sub WriteSection(ByVal shName As String, ByVal sec As Object)
    Dim ws As Worksheet
    Dim rows As Collection, cols As Object
    Dim useVdom As Boolean, off As Long, nCol As Long, nRow As Long
    Dim arr() As Variant
    Dim k As Variant, o As Variant, rec As Object
    Dim r As Long, c As Long, v As String
    Dim rng As Range

    Set ws = GetOrCreateSheet(shName)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear

    Set rows = sec("rows")
    Set cols = sec("cols")
    useVdom = sec("hasVdom")
    If useVdom Then off = 2 Else off = 1
    nCol = off + cols.Count
    nRow = rows.Count + 1

    ReDim arr(1 To nRow, 1 To nCol)
    If useVdom Then arr(1, 1) = "VDOM"
    arr(1, off) = "edit"
    For Each k In cols.Keys
        arr(1, off + cols(k)) = k
    Next k

    r = 1
    For Each o In rows
        Set rec = o
        r = r + 1
        If useVdom Then arr(r, 1) = rec("__vdom")
        arr(r, off) = rec("__edit")
        For Each k In rec.Keys
            If cols.Exists(k) Then
                v = rec(k)
                If Len(v) > MAX_CELL_LEN Then v = Left$(v, MAX_CELL_LEN) & "...(省略)"
                arr(r, off + cols(k)) = v
            End If
        Next k
    Next o

    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nRow, nCol))
    rng.NumberFormat = "@"          ' "=" 始まりの値や数値を文字列のまま保持
    rng.Value = arr
    rng.VerticalAlignment = xlTop

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nCol))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    rng.Columns.AutoFit
    For c = 1 To nCol
        If ws.Columns(c).ColumnWidth > MAX_COL_WIDTH Then ws.Columns(c).ColumnWidth = MAX_COL_WIDTH
    Next c
    rng.WrapText = True
    rng.Rows.AutoFit
    rng.AutoFilter

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(2, off + 1).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

'--------------------------------------------------------------
' シート関連
'--------------------------------------------------------------
Private Function GetSettingSheet() As Worksheet
    Dim ws As Worksheet, defs As Variant, exc As Variant, i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTING_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set GetSettingSheet = ws
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = SETTING_SHEET
    ws.Range("A1").Value = "抽出セクション"
    ws.Range("B1").Value = "件数"
    ws.Range("D1").Value = "除外キー"
    ws.Range("F1").Value = "読込ファイル / 日時"

    defs = Array("system interface", "system zone", "router static", _
                 "firewall address", "firewall addrgrp", _
                 "firewall service custom", "firewall service group", _
                 "firewall vip", "firewall vipgrp", "firewall ippool", _
                 "firewall schedule recurring", "firewall policy", _
                 "vpn ipsec phase1-interface", "vpn ipsec phase2-interface", _
                 "user local", "user group")
    For i = 0 To UBound(defs)
        ws.Cells(i + 2, 1).Value = defs(i)
    Next i

    ' 秘密情報系はデフォルトで出力しない
    exc = Array("uuid", "passwd", "password", "psksecret", "private-key")
    For i = 0 To UBound(exc)
        ws.Cells(i + 2, 4).Value = exc(i)
    Next i

    ws.Range("A1:F1").Font.Bold = True
    ws.Columns("A").ColumnWidth = 32
    ws.Columns("B").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 40
    Set GetSettingSheet = ws
End Function

Private Function ReadList(ByVal ws As Worksheet, ByVal col As Long) As Object
    Dim d As Object, r As Long, lastR As Long, v As String
    Set d = CreateObject("Scripting.Dictionary")
    lastR = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    For r = 2 To lastR
        v = NormalizeSpaces(LCase$(Trim$(CStr(ws.Cells(r, col).Value))))
        If Len(v) > 0 Then
            If Not d.Exists(v) Then d.Add v, r
        End If
    Next r
    Set ReadList = d
End Function

Private Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function SafeSheetName(ByVal s As String) As String
    Dim ch As Variant
    s = Replace(s, " ", "_")
    For Each ch In Array("\", "/", "?", "*", "[", "]", ":")
        s = Replace(s, ch, "")
    Next ch
    If Len(s) > 31 Then s = Left$(s, 31)
    SafeSheetName = s
End Function

'--------------------------------------------------------------
' 文字列ユーティリティ
'--------------------------------------------------------------
' set の値を分割（"..." 内の空白は保持、\" \\ のエスケープに対応）
Private Function Tokenize(ByVal s As String) As Collection
    Dim c As New Collection
    Dim i As Long, n As Long, ch As String, cur As String
    Dim inQ As Boolean, started As Boolean

    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If inQ Then
            If ch = "\" And i < n Then
                i = i + 1
                cur = cur & Mid$(s, i, 1)
            ElseIf ch = """" Then
                inQ = False
            Else
                cur = cur & ch
            End If
        Else
            If ch = """" Then
                inQ = True
                started = True
            ElseIf ch = " " Or ch = vbTab Or ch = vbLf Then
                If started Then
                    c.Add cur
                    cur = ""
                    started = False
                End If
            Else
                cur = cur & ch
                started = True
            End If
        End If
        i = i + 1
    Loop
    If started Then c.Add cur
    Set Tokenize = c
End Function

Private Function IsQuoteOpen(ByVal s As String) As Boolean
    Dim i As Long, n As Long, ch As String, inQ As Boolean
    If InStr(s, """") = 0 Then Exit Function
    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If inQ And ch = "\" Then
            i = i + 1               ' エスケープされた次の1文字を飛ばす
        ElseIf ch = """" Then
            inQ = Not inQ
        End If
        i = i + 1
    Loop
    IsQuoteOpen = inQ
End Function

Private Function JoinTokens(ByVal c As Collection, ByVal sep As String) As String
    Dim v As Variant, s As String, first As Boolean
    first = True
    For Each v In c
        If first Then
            s = v
            first = False
        Else
            s = s & sep & v
        End If
    Next v
    JoinTokens = s
End Function

Private Function FirstWord(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, " ")
    If p = 0 Then
        FirstWord = s
    Else
        FirstWord = Left$(s, p - 1)
    End If
End Function

Private Function NormalizeSpaces(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeSpaces = s
End Function

Private Function ReadTextUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2              ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadTextUtf8 = st.ReadText(-1)   ' adReadAll
    st.Close
End Function
