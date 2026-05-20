Attribute VB_Name = "CheckInCheck"
Option Explicit

'====================
' 配置区（只改这里）
'====================
Const SHEET_RELEASE As String = "release"
Const SHEET_RESULT As String = "result"
Const SHEET_AAA As String = "STリリース管理台帳"

' release 列
Const COL_KEYA As Long = 1 'A列
Const COL_KEYB As Long = 2 'B列

' aaa 列
Const COL_AAA_KEYC As Long = 4 'D列：B票番号
Const COL_AAA_NO As Long = 1 'A列：No
Const COL_AAA_STATUS_OUT As Long = 12 'L列：出庫ステータス
Const COL_AAA_STATUS_IN As Long = 20 'T列：入庫ステータス
Const COL_AAA_DAY As Long = 21 'U列：R判定予定日

' result 输出列（固定不建议改）
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

' target sheet 名
Dim TARGET_SHEETS As Variant

' target 列
Const COL_T_NO As Long = 1      '項番
Const COL_T_STUAT As Long = 3     '工程
Const COL_T_B_ID As Long = 5    'B票
Const COL_T_STATUS As Long = 17 '対応状況

Sub RunTool()
    Dim wbTool As Workbook, wbManage As Workbook, wbTarget As Workbook
    Dim wsRelease As Worksheet, wsResult As Worksheet, wsAAA As Worksheet
    Dim wsT As Worksheet
    
    Dim arrRelease, arrAAA
    Dim i As Long, r As Long
    Dim keyA, keyB, keyC
    Dim splitArr, item
    
    Dim dictKeyC As Object, dictKeyAB As Object
    Dim dictTarget As Object
    
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
    
    '读取
    Dim lastRowReleaseA As Long, lastRowReleaseB As Long, lastRowRelease As Long
    lastRowReleaseA = wsRelease.Cells(wsRelease.Rows.Count, COL_KEYA).End(xlUp).Row
    lastRowReleaseB = wsRelease.Cells(wsRelease.Rows.Count, COL_KEYB).End(xlUp).Row
    lastRowRelease = Application.WorksheetFunction.Max(lastRowReleaseA, lastRowReleaseB)
    arrRelease = wsRelease.Range(wsRelease.Cells(2, COL_KEYA), wsRelease.Cells(lastRowRelease, COL_KEYB)).Value2
    arrAAA = wsAAA.Range(wsAAA.Cells(2, 1), wsAAA.Cells(wsAAA.Cells(wsAAA.Rows.Count, COL_AAA_KEYC).End(xlUp).Row, COL_AAA_DAY)).Value2
    
    '构建 keyC 索引
    Set dictKeyC = CreateObject("Scripting.Dictionary")
    Set dictKeyAB = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(arrAAA)
        keyC = arrAAA(i, COL_AAA_KEYC)
        If keyC <> "" Then
            splitArr = Split(keyC, ",")
            For Each item In splitArr
                item = Trim(item)
                If item <> "" Then
                    If Not dictKeyC.exists(item) Then
                        dictKeyC(item) = Array(i)
                    Else
                        dictKeyC(item) = AppendIndex(dictKeyC(item), i)
                    End If
                End If
            Next
        End If
    Next
    
    '构建 target 索引
    Set dictTarget = CreateObject("Scripting.Dictionary")
    Dim arrT, lastRowT As Long, k
    
    For Each k In TARGET_SHEETS
        Set wsT = wbTarget.Sheets(k)
        lastRowT = wsT.Cells(wsT.Rows.Count, COL_T_STUAT).End(xlUp).Row
        If lastRowT >= 2 Then
            arrT = wsT.Range(wsT.Cells(2, COL_T_NO), wsT.Cells(lastRowT, COL_T_STATUS)).Value2
            For i = 1 To UBound(arrT)
                dictTarget(arrT(i, COL_T_STUAT) & arrT(i, COL_T_B_ID)) = Array("sheet" & k, arrT(i, COL_T_STATUS))
            Next
        End If
    Next
    
    '结果缓存
    Dim res()
    ReDim res(1 To 1000, 1 To 10)
    r = 0
    
    'keyA
    For i = 1 To UBound(arrRelease)
        keyA = CStr(arrRelease(i, 1))
        dictKeyAB("ST" & keyA) = 1
        If keyA <> "" Then
            If dictKeyC.exists(keyA) Then
                Dim rowsA, idx
                rowsA = dictKeyC(keyA)
                For Each idx In rowsA
                    splitArr = Split(arrAAA(idx, COL_AAA_KEYC), ",")
                    For Each item In splitArr
                        item = Trim(item)
                        r = r + 1
                        res(r, COL_RES_A) = "ST"
                        res(r, COL_RES_B) = keyA
                        
                        If InStr(item, "受入") > 0 Then
                            res(r, COL_RES_C) = "UAT"
                            res(r, COL_RES_D) = Replace(item, "受入", "")
                        Else
                            res(r, COL_RES_C) = "ST"
                            res(r, COL_RES_D) = item
                        End If
                        
                        res(r, COL_RES_E) = arrAAA(idx, COL_AAA_NO)
                        res(r, COL_RES_F) = arrAAA(idx, COL_AAA_STATUS_OUT)
                        res(r, COL_RES_G) = arrAAA(idx, COL_AAA_STATUS_IN)
                        res(r, COL_RES_H) = arrAAA(idx, COL_AAA_DAY)
                    Next
                Next
            Else
                r = r + 1
                res(r, COL_RES_A) = "ST"
                res(r, COL_RES_B) = keyA
                res(r, COL_RES_I) = "台帳に不存在"
            End If
        End If
    Next
    
    'keyB
    For i = 1 To UBound(arrRelease)
        keyB = arrRelease(i, 2)
        dictKeyAB("UAT" & keyB) = 1
        If keyB <> "" Then
            If dictKeyC.exists("受入" & keyB) Then
                Dim rowsB
                rowsB = dictKeyC("受入" & keyB)
                For Each idx In rowsB
                    splitArr = Split(arrAAA(idx, COL_AAA_KEYC), ",")
                    For Each item In splitArr
                        item = Trim(item)
                        r = r + 1
                        res(r, COL_RES_A) = "UAT"
                        res(r, COL_RES_B) = keyB
                        
                        If InStr(item, "受入") > 0 Then
                            res(r, COL_RES_C) = "UAT"
                            res(r, COL_RES_D) = Replace(item, "受入", "")
                        Else
                            res(r, COL_RES_C) = "ST"
                            res(r, COL_RES_D) = item
                        End If
                        
                        res(r, COL_RES_E) = arrAAA(idx, COL_AAA_NO)
                        res(r, COL_RES_F) = arrAAA(idx, COL_AAA_STATUS_OUT)
                        res(r, COL_RES_G) = arrAAA(idx, COL_AAA_STATUS_IN)
                        res(r, COL_RES_H) = arrAAA(idx, COL_AAA_DAY)
                    Next
                Next
            Else
                r = r + 1
                res(r, COL_RES_A) = "UAT"
                res(r, COL_RES_B) = keyB
                res(r, COL_RES_I) = "台帳に不存在"
            End If
        End If
    Next
    
    'target 查找
    For i = 1 To r
        If res(i, COL_RES_I) = "台帳に不存在" Then GoTo NEXTI
        Dim keyT
        keyT = res(i, COL_RES_C) & res(i, COL_RES_D)
        If dictTarget.exists(keyT) Then
            res(i, COL_RES_I) = dictTarget(keyT)(0)
            If Not dictKeyAB.exists(keyT) Then
                res(i, COL_RES_I) = res(i, COL_RES_I) & "（Not In Release List）"
            End If
            res(i, COL_RES_J) = dictTarget(keyT)(1)
        Else
            res(i, COL_RES_I) = "不是BBX的B票"
        End If
NEXTI:
    Next
    
    '写入
    wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(wsResult.Rows.Count, 10)).ClearContents
    If r > 0 Then
        wsResult.Cells(2, 1).Resize(r, 10).Value = res
        'wsResult.Rows("2:" & r + 1).Interior.Pattern = xlNone
        wsResult.Rows("2:" & wsResult.Rows.Count).Interior.Pattern = xlNone
        Call ColorByBlock(wsResult, 2, r + 1)
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    MsgBox "完成: " & r & " 件"
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
        
        ' 新??始
        If currentVal <> prevVal Then
            blockIndex = blockIndex + 1
            prevVal = currentVal
        End If
        
        ' 偶数? → 上色
        If blockIndex Mod 2 = 0 Then
            ws.Rows(i).Interior.Color = RGB(198, 239, 206) '浅?色
        Else
            ws.Rows(i).Interior.Pattern = xlNone
        End If
    Next i
End Sub
