Attribute VB_Name = "CheckInCheck"
Option Explicit

'====================
' 設定エリア（ここだけ変更）
'====================
Const SHEET_RELEASE As String = "release"
Const SHEET_RESULT As String = "result"
Const SHEET_AAA As String = "STリリース管理台帳"

' release シート列
Const COL_KEYA As Long = 1 'A列
Const COL_KEYB As Long = 2 'B列

' 台帳シート列
Const COL_AAA_KEYC As Long = 4 'D列：B票番号
Const COL_AAA_NO As Long = 1 'A列：No
Const COL_AAA_STATUS_OUT As Long = 12 'L列：出庫ステータス
Const COL_AAA_STATUS_IN As Long = 20 'T列：入庫ステータス
Const COL_AAA_DAY As Long = 21 'U列：R判定予定日

' result 出力列（固定のため変更非推奨）
Const COL_RES_A As Long = 1
Const COL_RES_B As Long = 2
Const COL_RES_C As Long = 3
Const COL_RES_D As Long = 4
Const COL_RES_E As Long = 5
Const COL_RES_F As Long = 6
Const COL_RES_G As Long = 7
Const COL_RES_H As Long = 8
Const COL_RES_I As Long = 9
Const COL_RES_J As Long = 10

' 参照対象シート名
Dim TARGET_SHEETS As Variant

' 参照対象シート列
Const COL_T_NO As Long = 1      '項番
Const COL_T_PROCESS As Long = 3     '工程
Const COL_T_B_ID As Long = 5    'B票
Const COL_T_STATUS As Long = 17 '対応状況

Sub RunTool()
    Dim wbTool As Workbook, wbManage As Workbook, wbTarget As Workbook
    Dim wsRelease As Worksheet, wsResult As Worksheet, wsAAA As Worksheet
    Dim wsT As Worksheet
    
    Dim releaseData, ledgerData
    Dim rowIndex As Long, resultCount As Long
    Dim releaseKeyST, releaseKeyUAT, ledgerKey
    Dim splitKeys, splitItem
    
    Dim ledgerKeyIndex As Object, releaseKeySet As Object
    Dim targetStatusIndex As Object
    
    TARGET_SHEETS = Array("対応中", "完了分", "品向T対応分")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set wbTool = ThisWorkbook
    Set wsRelease = wbTool.Sheets(SHEET_RELEASE)
    Set wsResult = wbTool.Sheets(SHEET_RESULT)
    
    Set wbManage = Workbooks("【ST】リリース管理台帳_V00200.xlsx")
    Set wsAAA = wbManage.Sheets(SHEET_AAA)
    
    Set wbTarget = Workbooks("ST不良対応状況報告書_BBX.xlsm")
    
    ' データ読み込み
    Dim lastRowReleaseA As Long, lastRowReleaseB As Long, lastRowRelease As Long
    lastRowReleaseA = wsRelease.Cells(wsRelease.Rows.Count, COL_KEYA).End(xlUp).Row
    lastRowReleaseB = wsRelease.Cells(wsRelease.Rows.Count, COL_KEYB).End(xlUp).Row
    lastRowRelease = Application.WorksheetFunction.Max(lastRowReleaseA, lastRowReleaseB)
    releaseData = wsRelease.Range(wsRelease.Cells(2, COL_KEYA), wsRelease.Cells(lastRowRelease, COL_KEYB)).Value2
    ledgerData = wsAAA.Range(wsAAA.Cells(2, 1), wsAAA.Cells(wsAAA.Cells(wsAAA.Rows.Count, COL_AAA_KEYC).End(xlUp).Row, COL_AAA_DAY)).Value2
    
    ' 台帳キー索引作成
    Set ledgerKeyIndex = CreateObject("Scripting.Dictionary")
    Set releaseKeySet = CreateObject("Scripting.Dictionary")
    
    For rowIndex = 1 To UBound(ledgerData)
        ledgerKey = ledgerData(rowIndex, COL_AAA_KEYC)
        If ledgerKey <> "" Then
            splitKeys = Split(ledgerKey, ",")
            For Each splitItem In splitKeys
                splitItem = Trim(splitItem)
                If splitItem <> "" Then
                    If Not ledgerKeyIndex.exists(splitItem) Then
                        ledgerKeyIndex(splitItem) = Array(rowIndex)
                    Else
                        ledgerKeyIndex(splitItem) = AppendIndex(ledgerKeyIndex(splitItem), rowIndex)
                    End If
                End If
            Next
        End If
    Next
    
    ' 対象ブック索引作成
    Set targetStatusIndex = CreateObject("Scripting.Dictionary")
    Dim targetData, lastRowTarget As Long, targetSheetName
    
    For Each targetSheetName In TARGET_SHEETS
        Set wsT = wbTarget.Sheets(targetSheetName)
        lastRowTarget = wsT.Cells(wsT.Rows.Count, COL_T_PROCESS).End(xlUp).Row
        If lastRowTarget >= 2 Then
            targetData = wsT.Range(wsT.Cells(2, COL_T_NO), wsT.Cells(lastRowTarget, COL_T_STATUS)).Value2
            For rowIndex = 1 To UBound(targetData)
                targetStatusIndex(targetData(rowIndex, COL_T_PROCESS) & targetData(rowIndex, COL_T_B_ID)) = Array("sheet" & targetSheetName, targetData(rowIndex, COL_T_STATUS))
            Next
        End If
    Next
    
    ' 結果バッファ初期化
    Dim resultData()
    ReDim resultData(1 To 1000, 1 To 10)
    resultCount = 0
    
    ' STキー処理
    For rowIndex = 1 To UBound(releaseData)
        releaseKeyST = CStr(releaseData(rowIndex, 1))
        releaseKeySet("ST" & releaseKeyST) = 1
        If releaseKeyST <> "" Then
            If ledgerKeyIndex.exists(releaseKeyST) Then
                Dim stMatchedRows, ledgerRow
                stMatchedRows = ledgerKeyIndex(releaseKeyST)
                For Each ledgerRow In stMatchedRows
                    splitKeys = Split(ledgerData(ledgerRow, COL_AAA_KEYC), ",")
                    For Each splitItem In splitKeys
                        splitItem = Trim(splitItem)
                        resultCount = resultCount + 1
                        resultData(resultCount, COL_RES_A) = "ST"
                        resultData(resultCount, COL_RES_B) = releaseKeyST
                        
                        If InStr(splitItem, "受入") > 0 Then
                            resultData(resultCount, COL_RES_C) = "UAT"
                            resultData(resultCount, COL_RES_D) = Replace(splitItem, "受入", "")
                        Else
                            resultData(resultCount, COL_RES_C) = "ST"
                            resultData(resultCount, COL_RES_D) = splitItem
                        End If
                        
                        resultData(resultCount, COL_RES_E) = ledgerData(ledgerRow, COL_AAA_NO)
                        resultData(resultCount, COL_RES_F) = ledgerData(ledgerRow, COL_AAA_STATUS_OUT)
                        resultData(resultCount, COL_RES_G) = ledgerData(ledgerRow, COL_AAA_STATUS_IN)
                        resultData(resultCount, COL_RES_H) = ledgerData(ledgerRow, COL_AAA_DAY)
                    Next
                Next
            Else
                resultCount = resultCount + 1
                resultData(resultCount, COL_RES_A) = "ST"
                resultData(resultCount, COL_RES_B) = releaseKeyST
                resultData(resultCount, COL_RES_I) = "台帳に不存在"
            End If
        End If
    Next
    
    ' UATキー処理
    For rowIndex = 1 To UBound(releaseData)
        releaseKeyUAT = releaseData(rowIndex, 2)
        releaseKeySet("UAT" & releaseKeyUAT) = 1
        If releaseKeyUAT <> "" Then
            If ledgerKeyIndex.exists("受入" & releaseKeyUAT) Then
                Dim uatMatchedRows
                uatMatchedRows = ledgerKeyIndex("受入" & releaseKeyUAT)
                For Each ledgerRow In uatMatchedRows
                    splitKeys = Split(ledgerData(ledgerRow, COL_AAA_KEYC), ",")
                    For Each splitItem In splitKeys
                        splitItem = Trim(splitItem)
                        resultCount = resultCount + 1
                        resultData(resultCount, COL_RES_A) = "UAT"
                        resultData(resultCount, COL_RES_B) = releaseKeyUAT
                        
                        If InStr(splitItem, "受入") > 0 Then
                            resultData(resultCount, COL_RES_C) = "UAT"
                            resultData(resultCount, COL_RES_D) = Replace(splitItem, "受入", "")
                        Else
                            resultData(resultCount, COL_RES_C) = "ST"
                            resultData(resultCount, COL_RES_D) = splitItem
                        End If
                        
                        resultData(resultCount, COL_RES_E) = ledgerData(ledgerRow, COL_AAA_NO)
                        resultData(resultCount, COL_RES_F) = ledgerData(ledgerRow, COL_AAA_STATUS_OUT)
                        resultData(resultCount, COL_RES_G) = ledgerData(ledgerRow, COL_AAA_STATUS_IN)
                        resultData(resultCount, COL_RES_H) = ledgerData(ledgerRow, COL_AAA_DAY)
                    Next
                Next
            Else
                resultCount = resultCount + 1
                resultData(resultCount, COL_RES_A) = "UAT"
                resultData(resultCount, COL_RES_B) = releaseKeyUAT
                resultData(resultCount, COL_RES_I) = "台帳に不存在"
            End If
        End If
    Next
    
    ' 対象ブック照合
    For rowIndex = 1 To resultCount
        If resultData(rowIndex, COL_RES_I) = "台帳に不存在" Then GoTo NEXT_RESULT
        Dim targetKey
        targetKey = resultData(rowIndex, COL_RES_C) & resultData(rowIndex, COL_RES_D)
        If targetStatusIndex.exists(targetKey) Then
            resultData(rowIndex, COL_RES_I) = targetStatusIndex(targetKey)(0)
            If Not releaseKeySet.exists(targetKey) Then
                resultData(rowIndex, COL_RES_I) = resultData(rowIndex, COL_RES_I) & "（Not In Release List）"
            End If
            resultData(rowIndex, COL_RES_J) = targetStatusIndex(targetKey)(1)
        Else
            resultData(rowIndex, COL_RES_I) = "BBXのB票ではない"
        End If
NEXT_RESULT:
    Next
    
    ' 結果書き込み
    wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(wsResult.Rows.Count, 10)).ClearContents
    If resultCount > 0 Then
        wsResult.Cells(2, 1).Resize(resultCount, 10).Value = resultData
        'wsResult.Rows("2:" & r + 1).Interior.Pattern = xlNone
        wsResult.Rows("2:" & wsResult.Rows.Count).Interior.Pattern = xlNone
        Call ColorByBlock(wsResult, 2, resultCount + 1)
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    MsgBox "完了: " & resultCount & " 件"
End Sub

Function AppendIndex(arr, val)
    Dim tmp(), i As Long
    ReDim tmp(0 To UBound(arr) + 1)
    For i = 0 To UBound(arr)
        tmp(i) = arr(i)
    Next
    tmp(UBound(tmp)) = val
    AppendIndex = tmp
End Function

Sub ColorByBlock(ws As Worksheet, startRow As Long, endRow As Long)
    Dim i As Long
    Dim currentVal As String, prevVal As String
    Dim blockIndex As Long
    
    blockIndex = 0
    prevVal = ""
    
    For i = startRow To endRow
        currentVal = CStr(ws.Cells(i, COL_RES_B).Value)
        
        ' 新しいブロック開始
        If currentVal <> prevVal Then
            blockIndex = blockIndex + 1
            prevVal = currentVal
        End If
        
        ' 偶数ブロックを着色
        If blockIndex Mod 2 = 0 Then
            ws.Rows(i).Interior.Color = RGB(198, 239, 206) '薄い緑色
        Else
            ws.Rows(i).Interior.Pattern = xlNone
        End If
    Next i
End Sub
