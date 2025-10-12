Option Explicit

Sub format_changeVer()

    Dim target As String '入力ファイルのフルパス
    Dim target_path As String '入力ファイルのパス
    Dim target_name As String '入力ファイルのファイル名
    Dim i As Long 'index
    Dim in_row As Long '入力ファイルの行カウンタ
    Dim param() As String 'Paramシートのパラメタ
    Dim last_row As Long '入力ファイルの最終行
    Dim wb As Workbook '出力ファイルのオブジェクト
    Dim out_row As Long '出力ファイルの行カウンタ
    Dim out_flg As Boolean '出力フラグ

    '入力ファイルを選択する(未選択はキャンセル扱い)
    target = Application.GetOpenFilename("Excel ブック,*.xlsx")
    If target = "False" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If

    '入力ファイルのパス、入力ファイルのファイル名を設定する
    target_name = Dir(target)
    target_path = Replace(target, target_name, "")

    'Paramシートのパラメタを動的配列へ格納する
    i = 0
    in_row = 2
    Do Until ThisWorkbook.Worksheets("Paramシート").Cells(in_row, 1).Value = ""
        ReDim Preserve param(i)
        param(i) = ThisWorkbook.Worksheets("Paramシート").Cells(in_row, 1).Value
        i = i + 1
        in_row = in_row + 1
    Loop

    '入力ファイルを開く
    Workbooks.Open target

    '入力ファイルの最終行を設定する
    last_row = Workbooks(target_name).Worksheets("MEMBERS").Cells(Rows.Count, 2).End(xlUp).Row

    '入力ファイルのヘッダを設定する
    Workbooks(target_name).Worksheets("MEMBERS").Range("Z1").Value = "absolute_number"
    Workbooks(target_name).Worksheets("MEMBERS").Range("AA1").Value = "FLG"

    '入力ファイルの明細を設定する
    in_row = 2
    Do Until in_row > last_row
        Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 26).Value = in_row - 1
        If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value = "" Then
            Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 27).Value = "1"
        Else
            Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 27).Value = "2"
        End If
        in_row = in_row + 1
    Loop

    '出力ファイルを作成する
    Set wb = Workbooks.Add

    '出力ファイルのヘッダを設定する
    On Error Resume Next
    wb.Worksheets(1).Name = "MEMBERS_" & Left(target_name, InStr(target_name, "00") - 1)
    On Error GoTo 0

    wb.Worksheets(1).Range("A1").Value = "Group"
    wb.Worksheets(1).Range("B1").Value = "Group Owners"
    wb.Worksheets(1).Range("C1").Value = "Group Recert"
    wb.Worksheets(1).Range("D1").Value = "Account"
    wb.Worksheets(1).Range("E1").Value = "Account Name"
    wb.Worksheets(1).Range("F1").Value = "Server Name"
    wb.Worksheets(1).Range("G1").Value = "Retain Access"
    wb.Worksheets(1).Range("H1").Value = "Remove Access"
    wb.Worksheets(1).Range("I1").Value = "Recertifier"
    wb.Worksheets(1).Range("J1").Value = "Application Name"
    wb.Worksheets(1).Range("K1").Value = "Service Account"
    wb.Worksheets(1).Range("L1").Value = "プライマリーオーナー"
    wb.Worksheets(1).Range("M1").Value = "再認定対象者"
    wb.Worksheets(1).Range("N1").Value = "XID/ID"
    wb.Worksheets(1).Range("O1").Value = "Function/所属"
    wb.Worksheets(1).Range("P1").Value = "Product/Tower"
    wb.Worksheets(1).Range("Q1").Value = "Sub Product"
    wb.Worksheets(1).Range("R1").Value = "Recertifier"
    wb.Worksheets(1).Range("S1").Value = "RSM(ID)"
    wb.Worksheets(1).Range("T1").Value = "RSM(Name)"
    wb.Worksheets(1).Range("Z1").Value = "absolute_number"
    wb.Worksheets(1).Range("AA1").Value = "FLG"
    wb.Worksheets(1).Range("AC1").Value = "BU"

    '出力ファイルの明細を設定する
    out_row = 2
    in_row = 2
    Do Until in_row > last_row
        If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value <> "" Then
            out_flg = False
            'Paramシートのパラメタ数分以下の処理を繰り返す
            i = 0
            Do Until i > UBound(param)
                If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value Like "*" & param(i) & "*" Then
                    out_flg = True
                End If
                i = i + 1
            Loop

            If out_flg = True Then
                wb.Worksheets(1).Cells(out_row, 1).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value
                wb.Worksheets(1).Cells(out_row, 2).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 2).Value
                wb.Worksheets(1).Cells(out_row, 3).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 3).Value
            End If
        End If

        '明細処理
        If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value = "" And _
           Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 2).Value <> "" And _
           out_flg = True Then

            If wb.Worksheets(1).Cells(out_row, 1).Value = "" Then
                wb.Worksheets(1).Cells(out_row, 1).Value = wb.Worksheets(1).Cells(out_row - 1, 1).Value
                wb.Worksheets(1).Cells(out_row, 2).Value = wb.Worksheets(1).Cells(out_row - 1, 2).Value
                wb.Worksheets(1).Cells(out_row, 3).Value = wb.Worksheets(1).Cells(out_row - 1, 3).Value
            End If

            wb.Worksheets(1).Cells(out_row, 4).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 2).Value
            wb.Worksheets(1).Cells(out_row, 5).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 3).Value
            wb.Worksheets(1).Cells(out_row, 26).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 26).Value
            wb.Worksheets(1).Cells(out_row, 27).Value = Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 27).Value
        End If

        in_row = in_row + 1
        If Workbooks(target_name).Worksheets("MEMBERS").Cells(in_row, 1).Value <> "" Then
            out_row = out_row + 1
        End If
    Loop

    '出力ファイルのフォーマット編集する
    wb.Worksheets(1).Range("A1:F1").Interior.Color = 16777062 '水色
    wb.Worksheets(1).Range("G1:I1").Interior.Color = 65535 '黄色
    wb.Worksheets(1).Range("J1:T1").Interior.Color = 5296274 '緑色
    wb.Worksheets(1).Range("A1:T" & out_row - 1).Font.Bold = True '太字
    wb.Worksheets(1).Range("A1:T" & out_row - 1).Borders.LineStyle = xlContinuous '格子
    wb.Worksheets(1).Columns("A:AC").AutoFit '自動調整
    wb.Worksheets(1).Columns("A:AC").AutoFilter 'フィルターをかける

    'ドロップダウンリストの追加
    With wb.Worksheets(1).Range("G2:G" & out_row)
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Retain"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    With wb.Worksheets(1).Range("H2:H" & out_row)
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Remove"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    '不要なシートを削除
    Application.DisplayAlerts = False
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(2).Delete
    Loop
    Application.DisplayAlerts = True

    '出力ファイルを保存する
    wb.SaveAs target_path & Replace(target_name, ".xlsx", "") & "_Member.xlsx"

    '出力ファイルを閉じる
    wb.Close

    '入力ファイルを保存する
    Workbooks(target_name).Save

    '入力ファイルを閉じる
    Workbooks(target_name).Close

    MsgBox "処理が完了しました。", vbInformation

End Sub
