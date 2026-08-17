Attribute VB_Name = "modOSTEngine"
'========================
'1) 标准模块：modOSTEngine
'========================
Option Explicit

Public Const TL_TITLE_COL As Long = 3      ' C
Public Const TL_ARTIST_COL As Long = 4     ' D
Public Const TL_ALBUM_COL As Long = 5      ' E
Public Const TL_NOTE_COL As Long = 6       ' F
Public Const TL_ID_COL As Long = 7         ' G，隐藏绑定编号

Public Const CL_NR_COL As Long = 1         ' A
Public Const CL_TITLE_COL As Long = 2      ' B
Public Const CL_ARTIST_COL As Long = 3     ' C
Public Const CL_ALBUM_COL As Long = 4      ' D
Public Const CL_NOTE_COL As Long = 5       ' E
Public Const CL_DISPLAY_COL As Long = 6    ' F
Public Const CL_ID_COL As Long = 7         ' G，隐藏绑定编号

Private mBusy As Boolean

Public Property Get OSTBusy() As Boolean
    OSTBusy = mBusy
End Property

Public Sub InitializeOSTSystem()
    On Error GoTo Fail

    mBusy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Tabelle1.Cells(1, TL_ID_COL).Value = "__OST_ID"
    Tabelle2.Cells(1, CL_ID_COL).Value = "__OST_ID"

    PrepareExistingCollectionIDs
    BindAllTimelineRows
    
    ' 清除历史遗留的重复 ID
    NormalizeTimelineIDsByTAAN
    
    RebuildCollectionByTimelineOrder
    RefreshTimelineValidation
    UpdateOSTStatistics

    ' ---------- Workspace ----------
    Tabelle1.Columns(TL_ID_COL).Hidden = True
    Tabelle2.Columns(CL_ID_COL).Hidden = True
    Tabelle2.Columns(CL_DISPLAY_COL).Hidden = True
    
    SyncEpisodeStatsButton

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mBusy = False
    Exit Sub

Fail:
    MsgBox "初始化失败：" & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub FormatTimestampInput(ByVal Target As Range)
    Dim cell As Range
    Dim rawText As String
    Dim resultText As String
    Dim hoursPart As Long
    Dim minutesPart As Long
    Dim secondsPart As Long

    If Target Is Nothing Then Exit Sub

    ' 支持一次粘贴多个时间戳
    For Each cell In Intersect(Target, Tabelle1.Range("B2:B2000")).Cells
        rawText = Trim$(CStr(cell.Value2))

        ' 空白、已经含冒号、含非数字字符：保持不变
        If rawText <> "" _
           And InStr(rawText, ":") = 0 _
           And rawText Like String(Len(rawText), "#") Then

            resultText = ""

            Select Case Len(rawText)

                Case 1, 2
                    ' 45 → 0:45
                    minutesPart = CLng(rawText)

                    If minutesPart <= 59 Then
                        resultText = "0:" & Format$(minutesPart, "00")
                    End If

                Case 3
                    ' 045 → 0:45
                    ' 145 → 1:45
                    minutesPart = CLng(Right$(rawText, 2))
                    hoursPart = CLng(Left$(rawText, 1))

                    If minutesPart <= 59 Then
                        resultText = hoursPart & ":" & _
                                     Format$(minutesPart, "00")
                    End If

                Case 4
                    ' 1252 → 12:52
                    minutesPart = CLng(Right$(rawText, 2))
                    hoursPart = CLng(Left$(rawText, 2))

                    If minutesPart <= 59 Then
                        resultText = hoursPart & ":" & _
                                     Format$(minutesPart, "00")
                    End If

                Case 5
                    ' 10302 → 1:03:02
                    secondsPart = CLng(Right$(rawText, 2))
                    minutesPart = CLng(Mid$(rawText, 2, 2))
                    hoursPart = CLng(Left$(rawText, 1))

                    If minutesPart <= 59 And secondsPart <= 59 Then
                        resultText = hoursPart & ":" & _
                                     Format$(minutesPart, "00") & ":" & _
                                     Format$(secondsPart, "00")
                    End If

                Case 6
                    ' 100302 → 10:03:02
                    secondsPart = CLng(Right$(rawText, 2))
                    minutesPart = CLng(Mid$(rawText, 3, 2))
                    hoursPart = CLng(Left$(rawText, 2))

                    If minutesPart <= 59 And secondsPart <= 59 Then
                        resultText = hoursPart & ":" & _
                                     Format$(minutesPart, "00") & ":" & _
                                     Format$(secondsPart, "00")
                    End If

            End Select

            If resultText <> "" Then
                Application.EnableEvents = False
                cell.NumberFormat = "@"
                cell.Value = resultText
                Application.EnableEvents = True
            End If
        End If
    Next cell
End Sub

Private Function GetMaxEpisodeNumber() As Long
    Dim wsTimeline As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim epText As String
    Dim epNumber As Long
    Dim maxEP As Long

    Set wsTimeline = Tabelle1

    lastRow = wsTimeline.Cells( _
        wsTimeline.Rows.Count, "A").End(xlUp).Row

    maxEP = 0

    For i = 2 To lastRow
        epText = UCase$(Trim$(CStr(wsTimeline.Cells(i, 1).Value)))
        epText = Replace(epText, " ", "")

        If Left$(epText, 2) = "EP" Then
            If IsNumeric(Mid$(epText, 3)) Then

                epNumber = CLng(Mid$(epText, 3))

                If epNumber > maxEP Then
                    maxEP = epNumber
                End If

            End If
        End If
    Next i

    GetMaxEpisodeNumber = maxEP
End Function

Public Sub UpdateNoMatchCount()
    Dim wsTimeline As Worksheet
    Dim wsCollection As Worksheet

    Dim lastRowT As Long
    Dim lastRowC As Long

    Dim i As Long
    Dim noMatchCount As Long
    Dim identifiedCount As Long

    Dim noteText As String
    Dim noMatchID As String

    Dim uniqueNoMatchIDs As Object

    Set wsTimeline = Tabelle1
    Set wsCollection = Tabelle2
    Set uniqueNoMatchIDs = CreateObject("Scripting.Dictionary")

    lastRowT = Application.Max( _
        wsTimeline.Cells(wsTimeline.Rows.Count, "A").End(xlUp).Row, _
        wsTimeline.Cells(wsTimeline.Rows.Count, "B").End(xlUp).Row, _
        wsTimeline.Cells(wsTimeline.Rows.Count, "C").End(xlUp).Row, _
        wsTimeline.Cells(wsTimeline.Rows.Count, "F").End(xlUp).Row)

    ' --------------------------------------------------
    ' No Match现在只统计：
    '
    ' A有Episode
    ' B有Timestamp
    ' C没有Title
    ' F含No-Match-ID
    '
    ' 相同ID只算一次
    ' --------------------------------------------------
    For i = 2 To lastRowT

        If Trim$(CStr(wsTimeline.Cells(i, 1).Value)) <> "" And _
           Trim$(CStr(wsTimeline.Cells(i, 2).Value)) <> "" And _
           Trim$(CStr(wsTimeline.Cells(i, 3).Value)) = "" Then

            noteText = CStr(wsTimeline.Cells(i, 6).Value)
            noMatchID = ExtractNoMatchID(noteText)

            If noMatchID <> "" Then
                If Not uniqueNoMatchIDs.Exists(noMatchID) Then
                    uniqueNoMatchIDs.Add noMatchID, True
                End If
            End If

        End If

    Next i

    noMatchCount = uniqueNoMatchIDs.Count

    ' Match = 表2当前实际收录条目数量
    lastRowC = LastCollectionRow

    If lastRowC >= 2 Then
        identifiedCount = _
            Application.WorksheetFunction.CountA( _
                wsCollection.Range("B2:B" & lastRowC))
    Else
        identifiedCount = 0
    End If

    wsCollection.Range("B1").Value = _
        "Title (Match: " & identifiedCount & _
        ", No Match: " & noMatchCount & ")"
End Sub
Private Sub NormalizeNoMatchIDs()
    Dim wsTimeline As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Dim noteText As String
    Dim noMatchID As String

    Dim oldNumbers As Object
    Dim numberList() As Long

    Dim oldNumber As Long
    Dim newNumber As Long

    Dim x As Long
    Dim y As Long
    Dim tempNumber As Long

    Dim re As Object

    Set wsTimeline = Tabelle1
    Set oldNumbers = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")

    With re
        .Pattern = "No-Match-(\d+)"
        .IgnoreCase = True
        .Global = False
    End With

    lastRow = Application.Max( _
        wsTimeline.Cells(wsTimeline.Rows.Count, "A").End(xlUp).Row, _
        wsTimeline.Cells(wsTimeline.Rows.Count, "B").End(xlUp).Row, _
        wsTimeline.Cells(wsTimeline.Rows.Count, "C").End(xlUp).Row, _
        wsTimeline.Cells(wsTimeline.Rows.Count, "F").End(xlUp).Row)

    ' 收集当前仍未识别的ID
    For i = 2 To lastRow

        If Trim$(CStr(wsTimeline.Cells(i, 1).Value)) <> "" And _
           Trim$(CStr(wsTimeline.Cells(i, 2).Value)) <> "" And _
           Trim$(CStr(wsTimeline.Cells(i, 3).Value)) = "" Then

            noteText = CStr(wsTimeline.Cells(i, 6).Value)

            If re.Test(noteText) Then
                oldNumber = CLng( _
                    re.Execute(noteText)(0).SubMatches(0))

                If Not oldNumbers.Exists(CStr(oldNumber)) Then
                    oldNumbers.Add CStr(oldNumber), oldNumber
                End If
            End If

        End If

    Next i

    If oldNumbers.Count = 0 Then Exit Sub

    ReDim numberList(1 To oldNumbers.Count)

    For x = 1 To oldNumbers.Count
        numberList(x) = CLng(oldNumbers.Items()(x - 1))
    Next x

    ' 从小到大排序
    For x = LBound(numberList) To UBound(numberList) - 1
        For y = x + 1 To UBound(numberList)

            If numberList(y) < numberList(x) Then
                tempNumber = numberList(x)
                numberList(x) = numberList(y)
                numberList(y) = tempNumber
            End If

        Next y
    Next x

    ' 旧编号 -> 新连续编号
    oldNumbers.RemoveAll

    For x = LBound(numberList) To UBound(numberList)
        oldNumbers.Add CStr(numberList(x)), x
    Next x

    ' 更新所有仍未识别行
    For i = 2 To lastRow

        If Trim$(CStr(wsTimeline.Cells(i, 1).Value)) <> "" And _
           Trim$(CStr(wsTimeline.Cells(i, 2).Value)) <> "" And _
           Trim$(CStr(wsTimeline.Cells(i, 3).Value)) = "" Then

            noteText = CStr(wsTimeline.Cells(i, 6).Value)

            If re.Test(noteText) Then

                oldNumber = CLng( _
                    re.Execute(noteText)(0).SubMatches(0))

                newNumber = CLng( _
                    oldNumbers(CStr(oldNumber)))

                wsTimeline.Cells(i, 6).Value = _
                    re.Replace( _
                        noteText, _
                        "No-Match-" & newNumber)
            End If

        End If

    Next i
End Sub
Private Function ExtractNoMatchID(ByVal noteText As String) As String
    Dim re As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")

    With re
        .Pattern = "No-Match-\d+"
        .IgnoreCase = True
        .Global = False
    End With

    If re.Test(noteText) Then
        Set matches = re.Execute(noteText)

        ExtractNoMatchID = _
            "No-Match-" & _
            Mid$(matches(0).Value, _
                 InStrRev(matches(0).Value, "-") + 1)
    Else
        ExtractNoMatchID = ""
    End If
End Function
Public Sub HandleNoMatchWorkflow(ByVal Target As Range)
    Dim hit As Range
    Dim cell As Range
    Dim currentRow As Long

    Set hit = Intersect(Target, Tabelle1.Range("A2:F2000"))
    If hit Is Nothing Then Exit Sub

    ' 本次操作涉及的最下面一行
    currentRow = 0

    For Each cell In hit.Cells
        If cell.Row > currentRow Then
            currentRow = cell.Row
        End If
    Next cell

    On Error GoTo CleanExit

    Application.EnableEvents = False

    ' 1. 已经识别成功的旧 No-Match 行：
    '    C有Title时，F整格临时备注全部清掉
    ClearResolvedNoMatchNotes

    ' 2. 某个No-Match被解决以后，
    '    先把剩余编号重新压成连续编号
    NormalizeNoMatchIDs

    ' 3. 当前正在登记的这一行不分配ID；
    '    只处理它上方已经被跨过去的未识别行
    If currentRow > 2 Then
        AssignPendingNoMatchIDs currentRow
    End If

CleanExit:
    Application.EnableEvents = True
End Sub

Private Sub ClearResolvedNoMatchNotes()
    Dim wsTimeline As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim noteText As String

    Set wsTimeline = Tabelle1

    lastRow = Application.Max( _
        wsTimeline.Cells(wsTimeline.Rows.Count, "C").End(xlUp).Row, _
        wsTimeline.Cells(wsTimeline.Rows.Count, "F").End(xlUp).Row)

    For i = 2 To lastRow

        ' C已经有Title
        If Trim$(CStr(wsTimeline.Cells(i, 3).Value)) <> "" Then

            noteText = CStr(wsTimeline.Cells(i, 6).Value)

            ' 只有F里确实存在No-Match-ID时才清
            ' 普通已识别曲自己的Note绝对不动
            If ExtractNoMatchID(noteText) <> "" Then
                wsTimeline.Cells(i, 6).ClearContents
            End If

        End If

    Next i
End Sub

Private Sub AssignPendingNoMatchIDs(ByVal currentRow As Long)
    Dim wsTimeline As Worksheet
    Dim i As Long
    Dim nextNumber As Long

    Dim epText As String
    Dim timestampText As String
    Dim titleText As String
    Dim noteText As String

    Set wsTimeline = Tabelle1

    nextNumber = GetNextNoMatchNumber()

    ' 只检查当前正在登记行之前的记录
    For i = 2 To currentRow - 1

        epText = Trim$(CStr(wsTimeline.Cells(i, 1).Value))
        timestampText = Trim$(CStr(wsTimeline.Cells(i, 2).Value))
        titleText = Trim$(CStr(wsTimeline.Cells(i, 3).Value))
        noteText = Trim$(CStr(wsTimeline.Cells(i, 6).Value))

        If epText <> "" And _
           timestampText <> "" And _
           titleText = "" Then

            ' 已经有ID的不动
            If ExtractNoMatchID(noteText) = "" Then

                If noteText = "" Then
                    wsTimeline.Cells(i, 6).Value = _
                        "No-Match-" & nextNumber & " | "
                Else
                    ' 已经写过场景备注，则保留原备注
                    wsTimeline.Cells(i, 6).Value = _
                        "No-Match-" & nextNumber & " | " & noteText
                End If

                nextNumber = nextNumber + 1
            End If

        End If

    Next i
End Sub

Private Function GetNextNoMatchNumber() As Long
    Dim wsTimeline As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Dim noteText As String
    Dim noMatchID As String
    Dim numberPart As Long
    Dim maxNumber As Long

    Set wsTimeline = Tabelle1

    lastRow = wsTimeline.Cells( _
        wsTimeline.Rows.Count, "F").End(xlUp).Row

    maxNumber = 0

    For i = 2 To lastRow

        noteText = CStr(wsTimeline.Cells(i, 6).Value)
        noMatchID = ExtractNoMatchID(noteText)

        If noMatchID <> "" Then

            numberPart = CLng( _
                Mid$(noMatchID, Len("No-Match-") + 1))

            If numberPart > maxNumber Then
                maxNumber = numberPart
            End If

        End If

    Next i

    GetNextNoMatchNumber = maxNumber + 1
End Function

Public Sub UpdateUsageStats()
    Dim wsTimeline As Worksheet
    Dim wsCollection As Worksheet

    Dim episodeCounts As Object
    Dim totalCounts As Object
    Dim firstAppears As Object

    Dim lastRowT As Long
    Dim lastRowC As Long
    Dim oldLastStatsCol As Long

    Dim i As Long
    Dim ep As Long
    Dim maxEP As Long

    Dim itemID As String
    Dim epText As String
    Dim timestampText As String
    Dim key As String

    Dim totalFromTimeline As Long
    Dim totalFromEpisodes As Long
    Dim countValue As Long
    
    Dim epStatsWereHidden As Boolean
    
    Dim oldLastStatsRow As Long

    Set wsTimeline = Tabelle1
    Set wsCollection = Tabelle2

    Set episodeCounts = CreateObject("Scripting.Dictionary")
    Set totalCounts = CreateObject("Scripting.Dictionary")
    Set firstAppears = CreateObject("Scripting.Dictionary")
    
    ' 记住更新统计前EP区当前的显示状态
    If wsCollection.Cells(1, 10).Value <> "" Then
        epStatsWereHidden = wsCollection.Columns(10).Hidden
    Else
        ' 第一次生成EP统计时默认展开
        epStatsWereHidden = False
    End If

    ' --------------------------------------------------
    ' 当前表1里实际出现到第几集
    ' --------------------------------------------------
    maxEP = GetMaxEpisodeNumber()

    ' --------------------------------------------------
    ' 固定表头
    ' --------------------------------------------------
    wsCollection.Cells(1, 8).Value = "Total [Toggle EP stats]"
    wsCollection.Cells(1, 9).Value = "First Appears"

    ' --------------------------------------------------
    ' 清掉以前可能残留的旧 EP 表头
    ' 从 J 列开始
    ' --------------------------------------------------
    oldLastStatsCol = wsCollection.Cells(1, _
        wsCollection.Columns.Count).End(xlToLeft).Column

    If oldLastStatsCol >= 10 Then
        wsCollection.Range( _
            wsCollection.Cells(1, 10), _
            wsCollection.Cells(1, oldLastStatsCol)).ClearContents
    End If

    ' --------------------------------------------------
    ' 动态建立 EP1 ~ EP最大值
    ' J = EP1
    ' --------------------------------------------------
    For ep = 1 To maxEP
        wsCollection.Cells(1, 9 + ep).Value = "EP" & ep
    Next ep

    lastRowT = LastTimelineRow
    lastRowC = LastCollectionRow

    ' --------------------------------------------------
    ' 扫描表1
    ' --------------------------------------------------
    For i = 2 To lastRowT

        itemID = Trim$(CStr( _
            wsTimeline.Cells(i, TL_ID_COL).Value))

        If itemID <> "" And _
           Trim$(CStr( _
               wsTimeline.Cells(i, TL_TITLE_COL).Value)) <> "" And _
           Not IsSummaryTimelineRow(i) Then

            ' ---------- 全表总次数 ----------
            If totalCounts.Exists(itemID) Then
                totalCounts(itemID) = totalCounts(itemID) + 1
            Else
                totalCounts.Add itemID, 1
            End If

            ' ---------- 首次出现 ----------
            If Not firstAppears.Exists(itemID) Then

                epText = Trim$(CStr( _
                    wsTimeline.Cells(i, 1).Value))

                timestampText = Trim$(CStr( _
                    wsTimeline.Cells(i, 2).Value))

                If epText <> "" And timestampText <> "" Then
                    firstAppears.Add itemID, _
                        epText & " - " & timestampText
                End If

            End If

            ' ---------- Episode ----------
            epText = UCase$(Trim$(CStr( _
                wsTimeline.Cells(i, 1).Value)))

            epText = Replace(epText, " ", "")

            If Left$(epText, 2) = "EP" Then

                If IsNumeric(Mid$(epText, 3)) Then

                    ep = CLng(Mid$(epText, 3))

                    If ep >= 1 Then

                        key = itemID & "|" & CStr(ep)

                        If episodeCounts.Exists(key) Then
                            episodeCounts(key) = _
                                episodeCounts(key) + 1
                        Else
                            episodeCounts.Add key, 1
                        End If

                    End If
                End If
            End If
        End If
    Next i

    ' --------------------------------------------------
    ' 清除旧统计数据
    ' H开始一直清到以前用过的最后统计列
    ' --------------------------------------------------
    oldLastStatsCol = Application.Max( _
        oldLastStatsCol, _
        9 + maxEP)
    
    ' 统计区以前可能比当前Collection更长，
    ' 所以必须把旧尾巴也一起清掉
    oldLastStatsRow = Application.Max( _
        lastRowC, _
        wsCollection.Cells(wsCollection.Rows.Count, 8).End(xlUp).Row, _
        wsCollection.Cells(wsCollection.Rows.Count, 9).End(xlUp).Row)
    
    If oldLastStatsRow >= 2 Then
        wsCollection.Range( _
            wsCollection.Cells(2, 8), _
            wsCollection.Cells(oldLastStatsRow, oldLastStatsCol) _
        ).ClearContents
    End If

    ' --------------------------------------------------
    ' 写回表2
    ' --------------------------------------------------
    For i = 2 To lastRowC

        itemID = Trim$(CStr( _
            wsCollection.Cells(i, CL_ID_COL).Value))

        If itemID <> "" Then

            ' ---------- First Appears ----------
            If firstAppears.Exists(itemID) Then
                wsCollection.Cells(i, 9).Value = _
                    firstAppears(itemID)
            Else
                wsCollection.Cells(i, 9).ClearContents
            End If

            totalFromEpisodes = 0

            ' ---------- EP1 ~ EP最大值 ----------
            For ep = 1 To maxEP

                key = itemID & "|" & CStr(ep)

                If episodeCounts.Exists(key) Then

                    countValue = CLng(episodeCounts(key))

                    wsCollection.Cells(i, 9 + ep).Value = _
                        countValue

                    totalFromEpisodes = _
                        totalFromEpisodes + countValue

                Else
                    wsCollection.Cells(i, 9 + ep).ClearContents
                End If

            Next ep

            ' ---------- 全表总次数 ----------
            If totalCounts.Exists(itemID) Then
                totalFromTimeline = _
                    CLng(totalCounts(itemID))
            Else
                totalFromTimeline = 0
            End If

            ' ---------- 双重统计核对 ----------
            If totalFromTimeline = totalFromEpisodes Then

                If totalFromTimeline > 0 Then
                    wsCollection.Cells(i, 8).Value = _
                        totalFromTimeline
                Else
                    wsCollection.Cells(i, 8).ClearContents
                End If

            Else
                wsCollection.Cells(i, 8).Value = _
                    "Count Mismatch"
            End If

        End If
    Next i

    ' 固定区域自动列宽
    wsCollection.Range("H:I").EntireColumn.AutoFit
    
    ' 重新生成EP列后，恢复用户原来的展开/隐藏状态
    If maxEP > 0 Then
        wsCollection.Range( _
            wsCollection.Columns(10), _
            wsCollection.Columns(9 + maxEP) _
        ).EntireColumn.Hidden = epStatsWereHidden
    End If
    
    ' 只同步按钮文字，不改变EP列状态
    SyncEpisodeStatsButton

End Sub

Public Sub ToggleEpisodeStats()
    Dim wsCollection As Worksheet
    Dim maxEP As Long

    Set wsCollection = Tabelle2

    maxEP = GetMaxEpisodeNumber()
    If maxEP = 0 Then Exit Sub

    ' J列 = EP1
    If wsCollection.Columns(10).Hidden Then
        ShowEpisodeStats
    Else
        HideEpisodeStats
    End If
End Sub


Public Sub ShowEpisodeStats()
    Dim wsCollection As Worksheet
    Dim maxEP As Long
    Dim lastEPCol As Long

    Set wsCollection = Tabelle2

    maxEP = GetMaxEpisodeNumber()
    If maxEP = 0 Then Exit Sub

    lastEPCol = 9 + maxEP

    ' 展开 J列开始的 EP1 ~ 当前最大EP
    wsCollection.Range( _
        wsCollection.Columns(10), _
        wsCollection.Columns(lastEPCol) _
    ).EntireColumn.Hidden = False

    ' 当前已经展开，所以按钮显示下一步操作
    wsCollection.Buttons("btnEPStats").Caption = "Hide EP Stats"
End Sub


Public Sub HideEpisodeStats()
    Dim wsCollection As Worksheet
    Dim maxEP As Long
    Dim lastEPCol As Long

    Set wsCollection = Tabelle2

    maxEP = GetMaxEpisodeNumber()
    If maxEP = 0 Then Exit Sub

    lastEPCol = 9 + maxEP

    ' 隐藏 J列开始的 EP1 ~ 当前最大EP
    wsCollection.Range( _
        wsCollection.Columns(10), _
        wsCollection.Columns(lastEPCol) _
    ).EntireColumn.Hidden = True

    ' 当前已经隐藏，所以按钮显示下一步操作
    wsCollection.Buttons("btnEPStats").Caption = "Show EP Stats"
End Sub

Public Sub SyncEpisodeStatsButton()
    Dim wsCollection As Worksheet
    Dim maxEP As Long

    Set wsCollection = Tabelle2

    maxEP = GetMaxEpisodeNumber()

    If maxEP = 0 Then
        wsCollection.Buttons("btnEPStats").Caption = "EP Stats"

    ElseIf wsCollection.Columns(10).Hidden Then
        wsCollection.Buttons("btnEPStats").Caption = "Show EP Stats"

    Else
        wsCollection.Buttons("btnEPStats").Caption = "Hide EP Stats"
    End If
End Sub

Public Sub UpdateOSTStatistics()
    UpdateNoMatchCount
    UpdateUsageStats
End Sub

Public Sub HandleTimelineChange(ByVal Target As Range)
    Dim hit As Range
    Dim r As Long
    Dim collectionRow As Long

    If mBusy Then Exit Sub

    Set hit = Intersect(Target, Tabelle1.Range("C2:F2000"))
    If hit Is Nothing Then Exit Sub

    On Error GoTo Fail

    mBusy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' 只有单格修改标题时，才判断是否通过下拉条目选取
    If Target.CountLarge = 1 And Target.Column = TL_TITLE_COL Then
        r = Target.Row
        collectionRow = CollectionRowFromDisplay(CStr(Target.Value))

        If collectionRow > 0 Then
            Tabelle1.Cells(r, TL_TITLE_COL).Value = Tabelle2.Cells(collectionRow, CL_TITLE_COL).Value
            Tabelle1.Cells(r, TL_ARTIST_COL).Value = Tabelle2.Cells(collectionRow, CL_ARTIST_COL).Value
            Tabelle1.Cells(r, TL_ALBUM_COL).Value = Tabelle2.Cells(collectionRow, CL_ALBUM_COL).Value
            Tabelle1.Cells(r, TL_NOTE_COL).Value = Tabelle2.Cells(collectionRow, CL_NOTE_COL).Value
            Tabelle1.Cells(r, TL_ID_COL).Value = Tabelle2.Cells(collectionRow, CL_ID_COL).Value
        End If
    End If

    BindChangedTimelineRows hit
    
    ' TAAN 完全相同的登记行必须统一为最早 ID
    NormalizeTimelineIDsByTAAN
    
    RebuildCollectionByTimelineOrder
    RefreshTimelineValidation
    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mBusy = False
    Exit Sub

Fail:
    MsgBox "表1同步失败：" & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub HandleCollectionChange(ByVal Target As Range)
    Dim hit As Range
    Dim r As Long
    Dim itemID As String
    Dim newTitle As String
    Dim i As Long
    Dim lastRowT As Long

    If mBusy Then Exit Sub

    Set hit = Intersect(Target, Tabelle2.Range("B2:E500"))
    If hit Is Nothing Then Exit Sub

    On Error GoTo Fail

    mBusy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' 多格操作、整行清空或粘贴时不再弹出意义不明的提示
    ' 直接按照表1的真实数据重新整理表2
    If Target.CountLarge <> 1 Then
        NormalizeTimelineIDsByTAAN
        RebuildCollectionByTimelineOrder
        RefreshTimelineValidation
        GoTo CleanExit
    End If

    r = Target.Row
    itemID = Trim$(CStr(Tabelle2.Cells(r, CL_ID_COL).Value))

    If itemID = "" Then
        NormalizeTimelineIDsByTAAN
        RebuildCollectionByTimelineOrder
        RefreshTimelineValidation
        GoTo CleanExit
    End If

    ' Titel 不能清空
    If Target.Column = CL_TITLE_COL Then
        newTitle = Trim$(CStr(Target.Value))

        If newTitle = "" Then
            Application.Undo
            MsgBox "Titel 不能单独清空。Artist、Album 和 Note 可以清空。", _
                   vbExclamation
            GoTo CleanExit
        End If
    End If

    lastRowT = LastTimelineRow

    ' 只更新原本绑定到这一条 ID 的登记行
    For i = 2 To lastRowT
        If CStr(Tabelle1.Cells(i, TL_ID_COL).Value) = itemID Then
            Tabelle1.Cells(i, TL_TITLE_COL).Value = _
                Tabelle2.Cells(r, CL_TITLE_COL).Value

            Tabelle1.Cells(i, TL_ARTIST_COL).Value = _
                Tabelle2.Cells(r, CL_ARTIST_COL).Value

            Tabelle1.Cells(i, TL_ALBUM_COL).Value = _
                Tabelle2.Cells(r, CL_ALBUM_COL).Value

            Tabelle1.Cells(i, TL_NOTE_COL).Value = _
                Tabelle2.Cells(r, CL_NOTE_COL).Value
        End If
    Next i

    ' 修改完成后，先按 TAAN 合并，再重新紧凑排列表2
    NormalizeTimelineIDsByTAAN
    RebuildCollectionByTimelineOrder
    RefreshTimelineValidation
    UpdateOSTStatistics

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mBusy = False
    Exit Sub

Fail:
    MsgBox "表2同步失败：" & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub PrepareExistingCollectionIDs()
    Dim i As Long
    Dim lastRowC As Long

    lastRowC = LastCollectionRow

    For i = 2 To lastRowC
        If Trim$(CStr(Tabelle2.Cells(i, CL_TITLE_COL).Value)) <> "" Then
            If Trim$(CStr(Tabelle2.Cells(i, CL_ID_COL).Value)) = "" Then
                Tabelle2.Cells(i, CL_ID_COL).Value = NewOSTID()
            End If
        End If
    Next i
End Sub

Private Sub BindAllTimelineRows()
    Dim i As Long
    Dim lastRowT As Long

    lastRowT = LastTimelineRow

    For i = 2 To lastRowT
        BindOneTimelineRow i
    Next i
End Sub

Private Sub BindChangedTimelineRows(ByVal hit As Range)
    Dim rowsDone As Object
    Dim c As Range
    Dim k As String

    Set rowsDone = CreateObject("Scripting.Dictionary")

    For Each c In hit.Cells
        k = CStr(c.Row)
        If Not rowsDone.Exists(k) Then
            rowsDone.Add k, True
            BindOneTimelineRow c.Row
        End If
    Next c
End Sub

Private Sub BindOneTimelineRow(ByVal r As Long)
    Dim titleText As String
    Dim currentID As String
    Dim collectionRow As Long

    titleText = Trim$(CStr(Tabelle1.Cells(r, TL_TITLE_COL).Value))
    currentID = Trim$(CStr(Tabelle1.Cells(r, TL_ID_COL).Value))

    If titleText = "" Then
        Tabelle1.Cells(r, TL_ID_COL).ClearContents
        Exit Sub
    End If

    ' 当前隐藏编号仍存在，且四项仍完全一致：继续绑定
    If currentID <> "" Then
        collectionRow = FindCollectionRowByID(currentID)

        If collectionRow > 0 Then
            If SameTAAN( _
                Tabelle1.Cells(r, TL_TITLE_COL).Value, _
                Tabelle1.Cells(r, TL_ARTIST_COL).Value, _
                Tabelle1.Cells(r, TL_ALBUM_COL).Value, _
                Tabelle1.Cells(r, TL_NOTE_COL).Value, _
                Tabelle2.Cells(collectionRow, CL_TITLE_COL).Value, _
                Tabelle2.Cells(collectionRow, CL_ARTIST_COL).Value, _
                Tabelle2.Cells(collectionRow, CL_ALBUM_COL).Value, _
                Tabelle2.Cells(collectionRow, CL_NOTE_COL).Value) Then
                Exit Sub
            End If
        End If
    End If

    ' 表1这一行被单独修改：先尝试绑定到已有完全相同条目
    collectionRow = FindCollectionRowByTAAN( _
        Tabelle1.Cells(r, TL_TITLE_COL).Value, _
        Tabelle1.Cells(r, TL_ARTIST_COL).Value, _
        Tabelle1.Cells(r, TL_ALBUM_COL).Value, _
        Tabelle1.Cells(r, TL_NOTE_COL).Value)

    If collectionRow > 0 Then
        Tabelle1.Cells(r, TL_ID_COL).Value = Tabelle2.Cells(collectionRow, CL_ID_COL).Value
    Else
        ' 找不到完全相同条目：只给当前登记行创建新编号
        Tabelle1.Cells(r, TL_ID_COL).Value = NewOSTID()
    End If
End Sub

Public Sub RebuildCollectionByTimelineOrder()
    Dim seen As Object
    Dim ids As Collection
    Dim firstRows As Collection
    Dim i As Long
    Dim lastRowT As Long
    Dim lastRowC As Long
    Dim itemID As String
    Dim outRow As Long
    Dim srcRow As Long

    Set seen = CreateObject("Scripting.Dictionary")
    Set ids = New Collection
    Set firstRows = New Collection

    lastRowT = LastTimelineRow

    ' 按表1首次出现顺序收集隐藏编号
    For i = 2 To lastRowT
        If Trim$(CStr(Tabelle1.Cells(i, TL_TITLE_COL).Value)) <> "" Then
            itemID = Trim$(CStr(Tabelle1.Cells(i, TL_ID_COL).Value))

            If itemID = "" Then
                itemID = NewOSTID()
                Tabelle1.Cells(i, TL_ID_COL).Value = itemID
            End If

            If Not seen.Exists(itemID) Then
                seen.Add itemID, True
                ids.Add itemID
                firstRows.Add i
            End If
        End If
    Next i

    lastRowC = Application.Max(LastCollectionRow, 2)
    Tabelle2.Range("A2:G" & Application.Max(lastRowC + 50, 600)).ClearContents

    outRow = 2

    For i = 1 To ids.Count
        itemID = CStr(ids(i))
        srcRow = CLng(firstRows(i))

        Tabelle2.Cells(outRow, CL_NR_COL).Value = outRow - 1
        Tabelle2.Cells(outRow, CL_TITLE_COL).Value = Tabelle1.Cells(srcRow, TL_TITLE_COL).Value
        Tabelle2.Cells(outRow, CL_ARTIST_COL).Value = Tabelle1.Cells(srcRow, TL_ARTIST_COL).Value
        Tabelle2.Cells(outRow, CL_ALBUM_COL).Value = Tabelle1.Cells(srcRow, TL_ALBUM_COL).Value
        Tabelle2.Cells(outRow, CL_NOTE_COL).Value = Tabelle1.Cells(srcRow, TL_NOTE_COL).Value

        ' 前置符号用于避免输入普通标题时被 Excel 自动匹配为下拉项
        Tabelle2.Cells(outRow, CL_DISPLAY_COL).Value = _
            (outRow - 1) & ". " & _
            Tabelle1.Cells(srcRow, TL_TITLE_COL).Value

        Tabelle2.Cells(outRow, CL_ID_COL).Value = itemID
        outRow = outRow + 1
    Next i

    Tabelle2.Range("A:E").EntireColumn.AutoFit
    Tabelle2.Range("H:I").EntireColumn.AutoFit
    
    Tabelle2.Columns(CL_DISPLAY_COL).Hidden = True
    Tabelle2.Columns(CL_ID_COL).Hidden = True
End Sub

Public Sub RefreshTimelineValidation()
    Dim lastRowC As Long
    Dim formulaText As String
    Dim safeName As String

    lastRowC = Tabelle2.Cells(Tabelle2.Rows.Count, CL_DISPLAY_COL).End(xlUp).Row
    If lastRowC < 2 Then lastRowC = 2

    safeName = Replace(Tabelle2.Name, "'", "''")
    formulaText = "='" & safeName & "'!$F$2:$F$" & lastRowC

    With Tabelle1.Range("C2:C2000").Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=formulaText
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = False
    End With

    Tabelle1.Columns(TL_ID_COL).Hidden = True
End Sub

Private Function CollectionRowFromDisplay(ByVal displayText As String) As Long
    Dim dotPos As Long
    Dim nrText As String
    Dim nr As Long
    Dim collectionRow As Long

    displayText = Trim$(displayText)

    ' 下拉名格式：序号. Titel
    dotPos = InStr(1, displayText, ". ", vbBinaryCompare)
    If dotPos <= 1 Then Exit Function

    nrText = Left$(displayText, dotPos - 1)

    ' 点号前必须全部是数字，普通曲名不应被当成下拉项
    If Not IsNumeric(nrText) Then Exit Function

    nr = CLng(nrText)
    If nr <= 0 Then Exit Function

    collectionRow = nr + 1

    ' 再验证表2 A列中的实际序号
    If CLng(Val(Tabelle2.Cells(collectionRow, CL_NR_COL).Value)) = nr Then
        CollectionRowFromDisplay = collectionRow
    End If
End Function

Private Function FindCollectionRowByID(ByVal itemID As String) As Long
    Dim found As Range

    If itemID = "" Then Exit Function

    Set found = Tabelle2.Columns(CL_ID_COL).Find( _
                    What:=itemID, _
                    After:=Tabelle2.Cells(1, CL_ID_COL), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=True)

    If Not found Is Nothing Then FindCollectionRowByID = found.Row
End Function

Private Function FindCollectionRowByTAAN( _
    ByVal titleText As Variant, _
    ByVal artistText As Variant, _
    ByVal albumText As Variant, _
    ByVal noteText As Variant) As Long

    Dim i As Long
    Dim lastRowC As Long

    lastRowC = LastCollectionRow

    For i = 2 To lastRowC
        If SameTAAN( _
            titleText, artistText, albumText, noteText, _
            Tabelle2.Cells(i, CL_TITLE_COL).Value, _
            Tabelle2.Cells(i, CL_ARTIST_COL).Value, _
            Tabelle2.Cells(i, CL_ALBUM_COL).Value, _
            Tabelle2.Cells(i, CL_NOTE_COL).Value) Then

            FindCollectionRowByTAAN = i
            Exit Function
        End If
    Next i
End Function

Private Function SameTAAN( _
    ByVal t1 As Variant, ByVal a1 As Variant, ByVal al1 As Variant, ByVal n1 As Variant, _
    ByVal t2 As Variant, ByVal a2 As Variant, ByVal al2 As Variant, ByVal n2 As Variant) As Boolean

    SameTAAN = _
        CleanText(t1) = CleanText(t2) And _
        CleanText(a1) = CleanText(a2) And _
        CleanText(al1) = CleanText(al2) And _
        CleanText(n1) = CleanText(n2)
End Function

Private Function CleanText(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        CleanText = ""
    Else
        CleanText = Trim$(CStr(v))
    End If
End Function

Private Function IsSummaryTimelineRow(ByVal rowNum As Long) As Boolean
    Dim timestampText As String

    timestampText = LCase$(Trim$(CStr(Tabelle1.Cells(rowNum, 2).Value)))

    Select Case timestampText
        Case "opening theme", "ending theme"
            IsSummaryTimelineRow = True

        Case Else
            IsSummaryTimelineRow = False
    End Select
End Function
Private Function LastTimelineRow() As Long
    LastTimelineRow = Application.Max( _
        Tabelle1.Cells(Tabelle1.Rows.Count, TL_TITLE_COL).End(xlUp).Row, _
        Tabelle1.Cells(Tabelle1.Rows.Count, TL_ID_COL).End(xlUp).Row, _
        2)
End Function

Private Function LastCollectionRow() As Long
    LastCollectionRow = Application.Max( _
        Tabelle2.Cells(Tabelle2.Rows.Count, CL_TITLE_COL).End(xlUp).Row, _
        Tabelle2.Cells(Tabelle2.Rows.Count, CL_ID_COL).End(xlUp).Row, _
        2)
End Function

Private Function NewOSTID() As String
    Randomize
    NewOSTID = Format$(Now, "yyyymmddhhnnss") & "_" & _
               Format$(CLng(Timer * 1000), "000000000") & "_" & _
               Format$(CLng(Rnd() * 1000000), "000000")
End Function

Public Sub DeleteSelectedCollectionEntry()
Attribute DeleteSelectedCollectionEntry.VB_ProcData.VB_Invoke_Func = "D\n14"
    Dim selectedRow As Long
    Dim deletedID As String
    Dim deletedTitle As String

    Dim sameTitleRows As Collection
    Dim lastRowC As Long
    Dim i As Long

    Dim replacementRow As Long
    Dim answer As Variant
    Dim promptText As String

    If ActiveSheet.CodeName <> Tabelle2.CodeName Then
        MsgBox "请先在 OST Collection 中选中要删除的曲条目。", vbInformation
        Exit Sub
    End If

    selectedRow = ActiveCell.Row

    If selectedRow < 2 Then Exit Sub

    deletedID = Trim$(CStr(Tabelle2.Cells(selectedRow, CL_ID_COL).Value))
    deletedTitle = CleanText(Tabelle2.Cells(selectedRow, CL_TITLE_COL).Value)

    If deletedID = "" Or deletedTitle = "" Then Exit Sub

    Set sameTitleRows = New Collection
    lastRowC = LastCollectionRow

    ' 查找其余同标题收录条
    For i = 2 To lastRowC
        If i <> selectedRow Then
            If CleanText(Tabelle2.Cells(i, CL_TITLE_COL).Value) = deletedTitle Then
                sameTitleRows.Add i
            End If
        End If
    Next i

    ' 没有其他同标题条目：不允许删除
    If sameTitleRows.Count = 0 Then
        MsgBox _
            "该曲目前只有这一条收录。" & vbCrLf & _
            "不能删除，否则表1中的相关登记将没有可替换信息。", _
            vbExclamation
        Exit Sub
    End If

    ' 只有一个其他版本：直接询问是否替换
    If sameTitleRows.Count = 1 Then
        replacementRow = CLng(sameTitleRows(1))

        If MsgBox( _
            "是否删除当前收录条，并把表1中属于它的所有登记改为：" & _
            vbCrLf & vbCrLf & _
            CollectionEntryText(replacementRow), _
            vbYesNo + vbQuestion, _
            "删除曲条目") <> vbYes Then

            Exit Sub
        End If

    Else
        ' 有多个其他版本：让用户输入要替换成的 Nr.
        promptText = _
            "请选择表1中的相关登记要改为哪一条。" & vbCrLf & _
            "请输入下面条目前面的 Nr.：" & vbCrLf & vbCrLf

        For i = 1 To sameTitleRows.Count
            promptText = promptText & _
                         CollectionEntryText(CLng(sameTitleRows(i))) & _
                         vbCrLf
        Next i

        answer = Application.InputBox( _
                     Prompt:=promptText, _
                     Title:="选择替换条目", _
                     Type:=1)

        If answer = False Then Exit Sub

        replacementRow = FindCollectionRowByNumber(CLng(answer))

        If replacementRow = 0 _
           Or replacementRow = selectedRow _
           Or CleanText(Tabelle2.Cells(replacementRow, CL_TITLE_COL).Value) <> deletedTitle Then

            MsgBox "输入的 Nr. 不是可用的同标题收录条。", vbExclamation
            Exit Sub
        End If
    End If

    ReplaceDeletedCollectionEntry deletedID, replacementRow
End Sub

Private Sub ReplaceDeletedCollectionEntry( _
    ByVal deletedID As String, _
    ByVal replacementRow As Long)

    Dim replacementID As String
    Dim lastRowT As Long
    Dim i As Long

    On Error GoTo Fail

    mBusy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    replacementID = CStr( _
        Tabelle2.Cells(replacementRow, CL_ID_COL).Value)

    lastRowT = LastTimelineRow

    ' 将被删除条目对应的表1登记全部换成所选条目
    For i = 2 To lastRowT
        If CStr(Tabelle1.Cells(i, TL_ID_COL).Value) = deletedID Then

            Tabelle1.Cells(i, TL_TITLE_COL).Value = _
                Tabelle2.Cells(replacementRow, CL_TITLE_COL).Value

            Tabelle1.Cells(i, TL_ARTIST_COL).Value = _
                Tabelle2.Cells(replacementRow, CL_ARTIST_COL).Value

            Tabelle1.Cells(i, TL_ALBUM_COL).Value = _
                Tabelle2.Cells(replacementRow, CL_ALBUM_COL).Value

            Tabelle1.Cells(i, TL_NOTE_COL).Value = _
                Tabelle2.Cells(replacementRow, CL_NOTE_COL).Value

            Tabelle1.Cells(i, TL_ID_COL).Value = replacementID
        End If
    Next i

    ' 自动合并、重排、连续编号
    NormalizeTimelineIDsByTAAN
    RebuildCollectionByTimelineOrder
    RefreshTimelineValidation
    UpdateOSTStatistics

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mBusy = False
    Exit Sub

Fail:
    MsgBox "删除曲条目失败：" & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function FindCollectionRowByNumber( _
    ByVal entryNumber As Long) As Long

    Dim found As Range

    Set found = Tabelle2.Columns(CL_NR_COL).Find( _
                    What:=entryNumber, _
                    After:=Tabelle2.Cells(1, CL_NR_COL), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext)

    If Not found Is Nothing Then
        FindCollectionRowByNumber = found.Row
    End If
End Function

Private Function CollectionEntryText( _
    ByVal collectionRow As Long) As String

    CollectionEntryText = _
        Tabelle2.Cells(collectionRow, CL_NR_COL).Value & ". " & _
        Tabelle2.Cells(collectionRow, CL_TITLE_COL).Value & _
        " | " & _
        Tabelle2.Cells(collectionRow, CL_ARTIST_COL).Value & _
        " | " & _
        Tabelle2.Cells(collectionRow, CL_ALBUM_COL).Value & _
        " | " & _
        Tabelle2.Cells(collectionRow, CL_NOTE_COL).Value
End Function

Public Sub RefreshOST()
    On Error GoTo Fail

    mBusy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    BindAllTimelineRows
    NormalizeTimelineIDsByTAAN
    RebuildCollectionByTimelineOrder
    RefreshTimelineValidation

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mBusy = False
    Exit Sub

Fail:
    MsgBox "刷新失败：" & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub NormalizeTimelineIDsByTAAN()
    Dim wsTimeline As Worksheet
    Dim firstIDByTAAN As Object
    Dim lastRowT As Long
    Dim i As Long

    Dim titleText As String
    Dim artistText As String
    Dim albumText As String
    Dim noteText As String
    Dim itemID As String
    Dim key As String

    Set wsTimeline = Tabelle1
    Set firstIDByTAAN = CreateObject("Scripting.Dictionary")

    lastRowT = LastTimelineRow

    For i = 2 To lastRowT
        titleText = CleanText(wsTimeline.Cells(i, TL_TITLE_COL).Value)

        If titleText = "" Then
            wsTimeline.Cells(i, TL_ID_COL).ClearContents
        Else
            artistText = CleanText(wsTimeline.Cells(i, TL_ARTIST_COL).Value)
            albumText = CleanText(wsTimeline.Cells(i, TL_ALBUM_COL).Value)
            noteText = CleanText(wsTimeline.Cells(i, TL_NOTE_COL).Value)

            ' 用完整 TAAN 作为真正的合并依据
            key = titleText & ChrW(30) & _
                  artistText & ChrW(30) & _
                  albumText & ChrW(30) & _
                  noteText

            itemID = Trim$(CStr(wsTimeline.Cells(i, TL_ID_COL).Value))

            If firstIDByTAAN.Exists(key) Then
                ' TAAN 完全一致：统一绑定到最早出现的 ID
                wsTimeline.Cells(i, TL_ID_COL).Value = firstIDByTAAN(key)
            Else
                ' 首次遇到该 TAAN：保留现有 ID，没有则新建
                If itemID = "" Then
                    itemID = NewOSTID()
                    wsTimeline.Cells(i, TL_ID_COL).Value = itemID
                End If

                firstIDByTAAN.Add key, itemID
            End If
        End If
    Next i
End Sub

Private Sub ReplaceTimelineBinding(ByVal oldID As String, ByVal replacementRow As Long)
    Dim newID As String
    Dim i As Long
    Dim lastRowT As Long

    On Error GoTo Fail

    mBusy = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    newID = CStr(Tabelle2.Cells(replacementRow, CL_ID_COL).Value)
    lastRowT = LastTimelineRow

    For i = 2 To lastRowT
        If CStr(Tabelle1.Cells(i, TL_ID_COL).Value) = oldID Then
            Tabelle1.Cells(i, TL_TITLE_COL).Value = Tabelle2.Cells(replacementRow, CL_TITLE_COL).Value
            Tabelle1.Cells(i, TL_ARTIST_COL).Value = Tabelle2.Cells(replacementRow, CL_ARTIST_COL).Value
            Tabelle1.Cells(i, TL_ALBUM_COL).Value = Tabelle2.Cells(replacementRow, CL_ALBUM_COL).Value
            Tabelle1.Cells(i, TL_NOTE_COL).Value = Tabelle2.Cells(replacementRow, CL_NOTE_COL).Value
            Tabelle1.Cells(i, TL_ID_COL).Value = newID
        End If
    Next i

    RebuildCollectionByTimelineOrder
    RefreshTimelineValidation

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mBusy = False
    Exit Sub

Fail:
    MsgBox "删除替换失败：" & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Function TAANLine(ByVal collectionRow As Long) As String
    TAANLine = Tabelle2.Cells(collectionRow, CL_NR_COL).Value & ". " & _
               Tabelle2.Cells(collectionRow, CL_TITLE_COL).Value & " | " & _
               Tabelle2.Cells(collectionRow, CL_ARTIST_COL).Value & " | " & _
               Tabelle2.Cells(collectionRow, CL_ALBUM_COL).Value & " | " & _
               Tabelle2.Cells(collectionRow, CL_NOTE_COL).Value
End Function

