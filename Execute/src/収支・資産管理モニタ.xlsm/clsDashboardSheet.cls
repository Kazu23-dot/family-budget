VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDashboardSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''*************************************************************************************
'''<summary>
''' 定数一覧(Dashboardシート)
'''</summary>
'''<option>
'''  Created：2022.05.01 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************

''Amazon売上CSVから取得したデータの格納用テーブル構造を定義
''******************************************************
'Private Const READ_AMAZON_EARNINGS_CSV_ENCODING As String = "Shift-JIS"
'Private Const READ_AMAZON_EARNINGS_CSV_DELIMITER As String = ","
'Private Const AMAZON_EARNINGS_CSV_HEADER_ROW_POS As Integer = 8
'Private Const EXTRACT_COL_AMAZON_EARNINGS_CSV_ELEMENT_CNT As Integer = 19
'
''Amazon売上CSVから抽出するカラムを定義
''******************************************************
'Private Const EXTRACT_AMAZON_EARNINGS_DATA_COL_DATE As Integer = 0       '<<=====日付
'Private Const EXTRACT_AMAZON_EARNINGS_DATA_COL_TRANSACTION As Integer = 2 '<<=====トランザクションの種類
'Private Const EXTRACT_AMAZON_EARNINGS_DATA_COL_ORDER_NO As Integer = 3   '<<=====注文番号
'Private Const EXTRACT_AMAZON_EARNINGS_DATA_COL_SKU As Integer = 4        '<<=====SKU
'Private Const EXTRACT_AMAZON_EARNINGS_DATA_COL_EXPLANATION As Integer = 5 '<<=====説明

'''*************************************************************************************
'''<summary>
''' 収入,支出,資産,負債データを収集(本体)
'''</summary>
'''<return></return>
'''<option>
'''  Created：2022.05.01 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Sub CollectIncomeExpenseAssetDebtData()
 '<<=============================================================>>
 '<<==[1]Amazon売上Csvファイルからデータ読取=======================>>
 '<<=============================================================>>
    '***指定された期間(年)を取得***
    Dim strTargetYearLst As VBA.Collection
    Set strTargetYearLst = New Collection
    Dim strYearStart As String
    Dim strYearEnd As String
    strYearStart = ThisWorkbook.Worksheets("ダッシュボード").Range("C2").Value
    strYearEnd = ThisWorkbook.Worksheets("ダッシュボード").Range("E2").Value
    If Val(strYearEnd) < Val(strYearStart) Then
        MsgBox "データ収集期間(年)が不正です。期間を指定しなおしてください。", vbYes
        Exit Sub
    End If
    Dim blnYearLoopFlag As Boolean
    Dim intTargetYear As Integer
    blnYearLoopFlag = True
    intTargetYear = Val(strYearStart)
    Do While blnYearLoopFlag
        If intTargetYear <= Val(strYearEnd) Then
            strTargetYearLst.Add (CStr(intTargetYear))
        Else
            blnYearLoopFlag = False
        End If
        intTargetYear = intTargetYear + 1
    Loop
    Dim strWorkBookLst As VBA.Collection
    Set strWorkBookLst = New Collection
    Dim strOpenFilePathFolder As String
    Dim strBufTmp As String
    strOpenFilePathFolder = ThisWorkbook.Path & "\InputData\"
    
    Dim objFSO As FileSystemObject: Set objFSO = New FileSystemObject
    Dim objFolder As Folder: Set objFolder = objFSO.GetFolder(strOpenFilePathFolder)
    Dim objFile As File
    
    For Each objFile In objFolder.Files
        strWorkBookLst.Add (objFile.Name)
    Next
    Dim strTargetWorkBookLst As VBA.Collection
    Set strTargetWorkBookLst = New Collection
    Dim varWorkBookNm As Variant
    Dim varTargetYear As Variant
    For Each varWorkBookNm In strWorkBookLst
        For Each varTargetYear In strTargetYearLst
            If InStr(varWorkBookNm, varTargetYear) Then
                strTargetWorkBookLst.Add (varWorkBookNm)
            End If
        Next
    Next

    '***データテーブルのデータをすべてクリアする***
    Dim objDeleteTable As ListObject
    Set objDeleteTable = ThisWorkbook.Sheets("DataSource").ListObjects("DataSource")
    If Not objDeleteTable.DataBodyRange Is Nothing Then
        objDeleteTable.DataBodyRange.Delete
    End If

    Dim varTargetWorkBook As Variant

    For Each varTargetWorkBook In strTargetWorkBookLst
        Dim varTargetYearArg As Variant

        For Each varTargetYearArg In strTargetYearLst
            If InStr(varTargetWorkBook, varTargetYearArg) Then
                '***処理対象年の設定(完了)***
                Exit For
            End If
        Next

        '***オープン先のファイルを指定***
        Dim strOpenFilePath As String
        strOpenFilePath = ThisWorkbook.Path & "\InputData\" & varTargetWorkBook

        Workbooks.Open Filename:=strOpenFilePath    '<<====データ取得先ブックオープン

        Dim objWriteWorkBook As Workbook
        Dim objReadWorkBook As Workbook

        Set objWriteWorkBook = ThisWorkbook
        Set objReadWorkBook = Workbooks(varTargetWorkBook)

        Dim blnLoopFlag As Boolean
        Dim intRowCnt As Integer
        Dim intReadLineCnt As Integer
        Dim intEmptyContinueCnt As Integer
        Dim strIncomeDataAry() As String
        blnLoopFlag = True
        intReadLineCnt = 1
        intEmptyContinueCnt = 0

        Do While blnLoopFlag
            Dim strCellValueTmp As String
            strCellValueTmp = objReadWorkBook.Worksheets(varTargetYearArg).Range("A" & intReadLineCnt).Value
            
            If InStr(strCellValueTmp, "#") > 0 Then
                If strCellValueTmp = "#給与収入#" Then
                    Call ReadExcelSheet(strIncomeDataAry(), varTargetYearArg, objReadWorkBook, intReadLineCnt, strCellValueTmp)

                    With objWriteWorkBook.Worksheets("DataSource").ListObjects("DataSource")
                        For intRowCnt = 0 To UBound(strIncomeDataAry(), 1)
                            If strIncomeDataAry(intRowCnt, 0) <> "" Then
                                .ListRows.Add (1)
                                .ListColumns("日付").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 2)
                                .ListColumns("メインカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 0)
                                .ListColumns("サブカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 1)
                                .ListColumns("収支額・資産負債額").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 3)
                            End If
                        Next
                    End With
                ElseIf strCellValueTmp = "#資産収入#" Then
                    Erase strIncomeDataAry
                    Call ReadExcelSheet(strIncomeDataAry(), varTargetYearArg, objReadWorkBook, intReadLineCnt, strCellValueTmp)
                    
                    With objWriteWorkBook.Worksheets("DataSource").ListObjects("DataSource")
                        For intRowCnt = 0 To UBound(strIncomeDataAry(), 1)
                            If strIncomeDataAry(intRowCnt, 0) <> "" Then
                                .ListRows.Add (1)
                                .ListColumns("日付").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 2)
                                .ListColumns("メインカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 0)
                                .ListColumns("サブカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 1)
                                .ListColumns("収支額・資産負債額").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 3)
                            End If
                        Next
                    End With
                ElseIf strCellValueTmp = "#事業収入#" Then
                    Erase strIncomeDataAry
                    Call ReadExcelSheet(strIncomeDataAry(), varTargetYearArg, objReadWorkBook, intReadLineCnt, strCellValueTmp)
                    
                    With objWriteWorkBook.Worksheets("DataSource").ListObjects("DataSource")
                        For intRowCnt = 0 To UBound(strIncomeDataAry(), 1)
                            If strIncomeDataAry(intRowCnt, 0) <> "" Then
                                .ListRows.Add (1)
                                .ListColumns("日付").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 2)
                                .ListColumns("メインカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 0)
                                .ListColumns("サブカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 1)
                                .ListColumns("収支額・資産負債額").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 3)
                            End If
                        Next
                    End With
                ElseIf strCellValueTmp = "#入金履歴#" Then
                    Erase strIncomeDataAry
                    Call ReadExcelSheet(strIncomeDataAry(), varTargetYearArg, objReadWorkBook, intReadLineCnt, strCellValueTmp)
                    
                    With objWriteWorkBook.Worksheets("DataSource").ListObjects("DataSource")
                        For intRowCnt = 0 To UBound(strIncomeDataAry(), 1)
                            If strIncomeDataAry(intRowCnt, 0) <> "" Then
                                .ListRows.Add (1)
                                .ListColumns("日付").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 2)
                                .ListColumns("メインカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 0)
                                .ListColumns("サブカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 1)
                                .ListColumns("収支額・資産負債額").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 3)
                            End If
                        Next
                    End With
                ElseIf strCellValueTmp = "#支出#" Then
                    Erase strIncomeDataAry
                    Call ReadExcelSheet(strIncomeDataAry(), varTargetYearArg, objReadWorkBook, intReadLineCnt, strCellValueTmp)
                    
                    With objWriteWorkBook.Worksheets("DataSource").ListObjects("DataSource")
                        For intRowCnt = 0 To UBound(strIncomeDataAry(), 1)
                            If strIncomeDataAry(intRowCnt, 0) <> "" Then
                                .ListRows.Add (1)
                                .ListColumns("日付").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 2)
                                .ListColumns("メインカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 0)
                                .ListColumns("サブカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 1)
                                .ListColumns("収支額・資産負債額").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 3)
                            End If
                        Next
                    End With
                ElseIf strCellValueTmp = "#資産推移#" Then
                    Erase strIncomeDataAry
                    Call ReadExcelSheet(strIncomeDataAry(), varTargetYearArg, objReadWorkBook, intReadLineCnt, strCellValueTmp)
                    
                    With objWriteWorkBook.Worksheets("DataSource").ListObjects("DataSource")
                        For intRowCnt = 0 To UBound(strIncomeDataAry(), 1)
                            If strIncomeDataAry(intRowCnt, 0) <> "" Then
                                .ListRows.Add (1)
                                .ListColumns("日付").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 2)
                                .ListColumns("メインカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 0)
                                .ListColumns("サブカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 1)
                                .ListColumns("収支額・資産負債額").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 3)
                            End If
                        Next
                    End With
                ElseIf strCellValueTmp = "#負債推移#" Then
                    Erase strIncomeDataAry
                    Call ReadExcelSheet(strIncomeDataAry(), varTargetYearArg, objReadWorkBook, intReadLineCnt, strCellValueTmp)
                    
                    With objWriteWorkBook.Worksheets("DataSource").ListObjects("DataSource")
                        For intRowCnt = 0 To UBound(strIncomeDataAry(), 1)
                            If strIncomeDataAry(intRowCnt, 0) <> "" Then
                                .ListRows.Add (1)
                                .ListColumns("日付").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 2)
                                .ListColumns("メインカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 0)
                                .ListColumns("サブカテゴリ").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 1)
                                .ListColumns("収支額・資産負債額").DataBodyRange(1) = strIncomeDataAry(intRowCnt, 3)
                            End If
                        Next
                    End With
                End If

                intEmptyContinueCnt = 0
            End If

            If strCellValueTmp = "" Then

                intEmptyContinueCnt = intEmptyContinueCnt + 1

                If intEmptyContinueCnt > 30 Then
                    blnLoopFlag = False
                End If
            End If
            intReadLineCnt = intReadLineCnt + 1
        Loop
        
        objReadWorkBook.Close '<<====データ取得先ブッククローズ
    Next
End Sub

'''*************************************************************************************
'''<summary>
''' 収支入力シートからデータ読取
'''</summary>
'''<param name="rStrReadExcelInputData()">ヘッダー行(何行目か)</param>
'''<param name="vObjTargetWorkBook">読取対象のブック</param>
'''<param name="vIntReadLineCnt">読取項目グループ名の行数</param>
'''<param name="vStrGroupNm">読取項目グループ名</param>
'''<return></return>
'''<option>
'''  Created：2022.04.29 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Sub ReadExcelSheet(ByRef rStrReadExcelInputData() As String, ByVal vStrTargetYear As String, ByVal vObjTargetWorkBook As Workbook, ByVal vIntReadLineCnt As Integer, ByVal vStrGroupNm As String)
  Dim blnLoopFlag As Boolean
  Dim intReadLineCnt As Integer
  ReDim rStrReadExcelInputData(23, 12)

  blnLoopFlag = True
  intReadLineCnt = 1
  
  Dim intDataCnt As Integer
  Dim intMonthCnt As Integer
  intDataCnt = 0
  intMonthCnt = 0

  If vStrGroupNm = "#給与収入#" Then
    ReDim rStrReadExcelInputData(23, 12)
  ElseIf vStrGroupNm = "#資産収入#" Then
    ReDim rStrReadExcelInputData(35, 12)
  ElseIf vStrGroupNm = "#事業収入#" Then
    ReDim rStrReadExcelInputData(0, 12)
  ElseIf vStrGroupNm = "#入金履歴#" Then
    ReDim rStrReadExcelInputData(23, 12)
  ElseIf vStrGroupNm = "#支出#" Then
    ReDim rStrReadExcelInputData(11, 12)
  ElseIf vStrGroupNm = "#資産推移#" Then
    ReDim rStrReadExcelInputData(47, 12)
  ElseIf vStrGroupNm = "#負債推移#" Then
    ReDim rStrReadExcelInputData(83, 12)
  End If

  Do While blnLoopFlag
    Dim strCellValueTmp As String
    Dim strInputData As String
    strCellValueTmp = vObjTargetWorkBook.Worksheets(vStrTargetYear).Range("B" & vIntReadLineCnt + intReadLineCnt).Value
    strInputData = ""

    '****************************************************************************************
    If vStrGroupNm = "#給与収入#" Then
        If strCellValueTmp = "総支給額" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value

                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "収入"
                    rStrReadExcelInputData(intDataCnt, 1) = "給与収入(額面)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "差引支給額" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value

                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "収入"
                    rStrReadExcelInputData(intDataCnt, 1) = "給与収入(手取)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "" Then
            blnLoopFlag = False
        End If
    '****************************************************************************************
    ElseIf vStrGroupNm = "#資産収入#" Then
        If strCellValueTmp = "国内株式(配当)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value

                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "収入"
                    rStrReadExcelInputData(intDataCnt, 1) = "国内株式(配当)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "外国株式(配当)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value

                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "収入"
                    rStrReadExcelInputData(intDataCnt, 1) = "外国株式(配当)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "売買損益" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value

                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "収入"
                    rStrReadExcelInputData(intDataCnt, 1) = "売買損益"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "" Then
            blnLoopFlag = False
        End If
    '****************************************************************************************
    ElseIf vStrGroupNm = "#事業収入#" Then
        If strCellValueTmp = "" Then
            blnLoopFlag = False
        End If
    '****************************************************************************************
    ElseIf vStrGroupNm = "#入金履歴#" Then
        If strCellValueTmp = "入金額" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "収入"
                    rStrReadExcelInputData(intDataCnt, 1) = "入金額"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "預金利息" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "収入"
                    rStrReadExcelInputData(intDataCnt, 1) = "預金利息"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "" Then
            blnLoopFlag = False
        End If
    '****************************************************************************************
    ElseIf vStrGroupNm = "#支出#" Then
        If strCellValueTmp = "合計支出額" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "支出"
                    rStrReadExcelInputData(intDataCnt, 1) = "支出総額"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "" Then
            blnLoopFlag = False
        End If
    '****************************************************************************************
    ElseIf vStrGroupNm = "#資産推移#" Then
        If strCellValueTmp = "預金" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "資産"
                    rStrReadExcelInputData(intDataCnt, 1) = "預金"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "株式(現物)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "資産"
                    rStrReadExcelInputData(intDataCnt, 1) = "株式(現物)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "投資信託" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "資産"
                    rStrReadExcelInputData(intDataCnt, 1) = "投資信託"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "確定拠出年金" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "資産"
                    rStrReadExcelInputData(intDataCnt, 1) = "確定拠出年金"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "" Then
            blnLoopFlag = False
        End If
    '****************************************************************************************
    ElseIf vStrGroupNm = "#負債推移#" Then
        If strCellValueTmp = "奨学金(第一種)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "負債"
                    rStrReadExcelInputData(intDataCnt, 1) = "奨学金(第一種)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "奨学金(第二種)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "負債"
                    rStrReadExcelInputData(intDataCnt, 1) = "奨学金(第二種)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "クレジット(PGU-Visa)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "負債"
                    rStrReadExcelInputData(intDataCnt, 1) = "クレジット(PGU-Visa)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "クレジット(楽天)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "負債"
                    rStrReadExcelInputData(intDataCnt, 1) = "クレジット(楽天)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "クレジット(Yahoo)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "負債"
                    rStrReadExcelInputData(intDataCnt, 1) = "クレジット(Yahoo)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "クレジット(dカード)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "負債"
                    rStrReadExcelInputData(intDataCnt, 1) = "クレジット(Yahoo)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "クレジット(auPay)" Then
            For intMonthCnt = 1 To 12
                '***入力データ***
                strInputData = vObjTargetWorkBook.Worksheets(vStrTargetYear).Cells(vIntReadLineCnt + intReadLineCnt, 3 + (intMonthCnt - 1)).Value
                If strInputData <> "" Then
                    rStrReadExcelInputData(intDataCnt, 0) = "負債"
                    rStrReadExcelInputData(intDataCnt, 1) = "クレジット(auPay)"
                    rStrReadExcelInputData(intDataCnt, 2) = vStrTargetYear & "/" & intMonthCnt
                    rStrReadExcelInputData(intDataCnt, 3) = strInputData

                    intDataCnt = intDataCnt + 1
                End If
            Next
        ElseIf strCellValueTmp = "" Then
            blnLoopFlag = False
        End If
    End If

    intReadLineCnt = intReadLineCnt + 1
  Loop
End Sub


