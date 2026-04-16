Option Explicit

'====================
' 配置区（只改这里）
'====================
Const SHEET_RELEASE As String = "release"
Const SHEET_RESULT As String = "result"
Const SHEET_AAA As String = "aaa"

' release 列
Const COL_KEYA As Long = 1 'A列
Const COL_KEYB As Long = 2 'B列

' aaa 列
Const COL_AAA_KEYC As Long = 1 'A列
Const COL_AAA_NO As Long = 2 'B列
Const COL_AAA_STATUS As Long = 3 'C列

' result 输出列（固定不建议改）
Const COL_RES_A As Long = 1
Const COL_RES_B As Long = 2
Const COL_RES_C As Long = 3
Const COL_RES_D As Long = 4
Const COL_RES_E As Long = 5
Const COL_RES_F As Long = 6
Const COL_RES_G As Long = 7
Const COL_RES_H As Long = 8

' target sheet 名
Dim TARGET_SHEETS As Variant

' target 列
Const COL_T_A As Long = 1
Const COL_T_B As Long = 2
Const COL_T_E As Long = 5
Const COL_T_F As Long = 6

Sub RunTool()
    Dim wbTool As Workbook, wbManage As Workbook, wbTarget As Workbook
    Dim wsRelease As Worksheet, wsResult As Worksheet, wsAAA As Worksheet
    Dim wsT As Worksheet
    
    Dim arrRelease, arrAAA
    Dim i As Long, r As Long
    Dim keyA, keyB, keyC
    Dim splitArr, item
    
    Dim dictKeyC As Object
    Dim dictTarget As Object
    
    TARGET_SHEETS = Array("対応中", "完了", "品質向上")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set wbTool = ThisWorkbook
    Set wsRelease = wbTool.Sheets(SHEET_RELEASE)
    Set wsResult = wbTool.Sheets(SHEET_RESULT)
    
    Set wbManage = Workbooks("manageFile.xlsx")
    Set wsAAA = wbManage.Sheets(SHEET_AAA)
    
    Set wbTarget = Workbooks("targetFile.xlsx")
    
    '读取
    arrRelease = wsRelease.Range(wsRelease.Cells(2, COL_KEYA), wsRelease.Cells(wsRelease.Cells(wsRelease.Rows.Count, COL_KEYA).End(xlUp).Row, COL_KEYB)).Value2
    arrAAA = wsAAA.Range(wsAAA.Cells(2, COL_AAA_KEYC), wsAAA.Cells(wsAAA.Cells(wsAAA.Rows.Count, COL_AAA_KEYC).End(xlUp).Row, COL_AAA_STATUS)).Value2
    
    '构建 keyC 索引
    Set dictKeyC = CreateObject("Scripting.Dictionary")
    
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
        lastRowT = wsT.Cells(wsT.Rows.Count, COL_T_B).End(xlUp).Row
        If lastRowT >= 2 Then
            arrT = wsT.Range(wsT.Cells(2, COL_T_A), wsT.Cells(lastRowT, COL_T_F)).Value2
            For i = 1 To UBound(arrT)
                dictTarget(arrT(i, COL_T_B) & arrT(i, COL_T_E)) = Array(arrT(i, COL_T_A), arrT(i, COL_T_F))
            Next
        End If
    Next
    
    '结果缓存
    Dim res()
    ReDim res(1 To 100000, 1 To 8)
    r = 0
    
    'keyA
    For i = 1 To UBound(arrRelease)
        keyA = arrRelease(i, 1)
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
                        res(r, COL_RES_F) = arrAAA(idx, COL_AAA_STATUS)
                    Next
                Next
            Else
                r = r + 1
                res(r, COL_RES_A) = "ST"
                res(r, COL_RES_B) = keyA
                res(r, COL_RES_G) = "台帳に不存在"
            End If
        End If
    Next
    
    'keyB
    For i = 1 To UBound(arrRelease)
        keyB = arrRelease(i, 2)
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
                        res(r, COL_RES_F) = arrAAA(idx, COL_AAA_STATUS)
                    Next
                Next
            Else
                r = r + 1
                res(r, COL_RES_A) = "UAT"
                res(r, COL_RES_B) = keyB
                res(r, COL_RES_G) = "台帳に不存在"
            End If
        End If
    Next
    
    'target 查找
    For i = 1 To r
        If res(i, COL_RES_G) = "台帳に不存在" Then GoTo NEXTI
        Dim keyT
        keyT = res(i, COL_RES_C) & res(i, COL_RES_D)
        If dictTarget.exists(keyT) Then
            res(i, COL_RES_G) = dictTarget(keyT)(0)
            res(i, COL_RES_H) = dictTarget(keyT)(1)
        Else
            res(i, COL_RES_G) = "不是BBX的B票"
        End If
NEXTI:
    Next
    
    '写入
    wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(wsResult.Rows.Count, 8)).ClearContents
    If r > 0 Then wsResult.Cells(2, 1).Resize(r, 8).Value = res
    
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
