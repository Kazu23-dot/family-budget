Attribute VB_Name = "mdlMain"

'''*************************************************************************************
'''<summary>
''' 収入,支出,資産,負債データを収集
'''</summary>
'''<return></return>
'''<option>
''' 2022.05.01 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Sub CollectIncomeExpenseAssetDebtData()
    Dim objDashboardSheet As New clsDashboardSheet
    objDashboardSheet.CollectIncomeExpenseAssetDebtData
End Sub

'''*************************************************************************************
'''<summary>
''' 期間コンボの選択肢を生成 ※現在年から過去40年分選択肢を生成
'''</summary>
'''<return></return>
'''<option>
''' 2022.05.05 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Sub CreateYearCmbSelection()
 '<<=============================================================>>
 '<<==[1]対象テーブル内のデータをクリア============================>>
 '<<=============================================================>>
    Dim objDeleteTable As ListObject
    Set objDeleteTable = ThisWorkbook.Sheets("Settings").ListObjects("SelectYearCmb")
    If Not objDeleteTable.DataBodyRange Is Nothing Then
        objDeleteTable.DataBodyRange.Delete
    End If

 '<<=============================================================>>
 '<<==[2]現在日時から年情報を取得し過去40年分の選択肢を生成==========>>
 '<<=============================================================>>
    Dim intYearTmp As Integer
    Dim strNowYear As String
    strNowYear = CStr(Year(Now))
    intYearTmp = strNowYear

    With ThisWorkbook.Sheets("Settings").ListObjects("SelectYearCmb")
        For intRowCnt = 1 To 40
            .ListRows.Add (1)
            .ListColumns("選択肢").DataBodyRange(1) = CStr(intYearTmp)
            intYearTmp = CInt(strNowYear) - intRowCnt
        Next
    End With
 '<<=============================================================>>
 '<<==[3]生成した年コンボボックス用のテーブルをソート================>>
 '<<=============================================================>>
    With ThisWorkbook.Sheets("Settings").ListObjects("SelectYearCmb")
        .Range.Sort key1:=.ListColumns("選択肢").Range, order1:=xlDescending, Header:=xlYes
    End With
End Sub

'''*************************************************************************************
'''<summary>
''' テスト実行用のメソッド
'''</summary>
'''<param name="vIntHeaderRowPos">ヘッダー行(何行目か)</param>
'''<param name="vStrEncoding">文字コード</param>
'''<return>行数</return>
'''<option>
'''  Created：2022.04.03 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Public Sub ExecuteTestMode()
    With ThisWorkbook.Sheets("DataSource").ListObjects("DataSource")
        .Range.Sort key1:=.ListColumns("メインカテゴリ").Range, order1:=xlDescending, _
                    key2:=.ListColumns("収支額・資産負債額").Range, order2:=xlDescending, Header:=xlYes
    End With
End Sub
