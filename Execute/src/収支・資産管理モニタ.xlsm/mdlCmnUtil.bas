Attribute VB_Name = "mdlCmnUtil"
Option Explicit

'''*************************************************************************************
'''<summary>
''' 連続したレコード数を取得
''' ワークシート,開始列,開始行を指定し処理実行
'''</summary>
'''<param name="vObjWorkSheet">対象ワークシート名</param>
'''<param name="vStrStartCol">開始列</param>
'''<param name="vIntStartRow">開始行</param>
'''<return>レコード数</return>
'''<option>
''' 2021.06.12 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Public Function GetRowCnt(vObjWorkSheet As Worksheet, vStrStartCol As String, vIntStartRow As Integer) As Integer
 Dim blnContinue As Boolean
 Dim intRowCnt As Integer
 
 '初期値設定
 blnContinue = True
 intRowCnt = 0
 
 '空データを発見するまで無限ループ
 Do While blnContinue
   If vObjWorkSheet.Range(vStrStartCol & (vIntStartRow + intRowCnt)).Value <> "" Then
     intRowCnt = intRowCnt + 1
   Else
     'カウント終了
     blnContinue = False
   End If
 Loop
   
 '戻り値を設定
 GetRowCnt = intRowCnt
End Function

'''*************************************************************************************
'''<summary>
''' 指定ファイルから行数を取得
'''</summary>
'''<param name="vStrFilePath"></param>
'''<return>行数</return>
'''<option>
''' 2021.06.12 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Public Function GetFileLineCount(vStrFilePath As String) As Long
    Dim objFS As New FileSystemObject
    Dim objTS As TextStream
    
    '引数のファイルが存在しない場合は処理を終了
    If (objFS.FileExists(vStrFilePath) = False) Then
        GetFileLineCount = -1
        Exit Function
    End If
    
    '追加モードで開く
    Set objTS = objFS.OpenTextFile(vStrFilePath, ForAppending)
    
    '戻り値を設定
    GetFileLineCount = objTS.Line - 1
End Function

'''*************************************************************************************
'''<summary>
''' 指定されたCSVからデータ取得
'''</summary>
'''<param name="vIntHeaderRowPos">ヘッダー行(何行目か)</param>
'''<param name="vStrEncoding">文字コード</param>
'''<return>行数</return>
'''<option>
'''  Created：2022.04.03 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Public Function ReadCsvFile(ByVal vIntHeaderRowPos As Integer, ByVal vStrEncoding As String, ByVal vStrDelimiter As String, _
                            ByRef rIntReadColCnt As Integer, ByRef rIntReadRowCnt As Integer)
  Dim blnRes As String
  Dim strAryCsvData() As String

  Dim strFolderPath As String
  Dim objMsgRslt As VbMsgBoxResult
  objMsgRslt = MsgBox(("CSVファイルのインポートを実行しますか？"), vbYesNo + vbQuestion, "インポート確認")

  'インポートを中止する場合
  If objMsgRslt = vbNo Then
     '処理終了
     Exit Function
  End If

  If Application.FileDialog(msoFileDialogFilePicker).Show = -1 Then
      'ファイル選択後にOKが押下された場合
      strFolderPath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
  Else
      'キャンセルが選択された場合
      MsgBox "処理を中止します", vbCritical
      Exit Function
  End If

  Dim intLineCnt As Integer
  intLineCnt = GetFileLineCount(strFolderPath)
  
  Dim intCommaCount As Integer
  intCommaCount = GetDelimiterCount(strFolderPath, vStrDelimiter)

  If (intLineCnt = -1) Or (intCommaCount = -1) Then
    'ファイルにデータが存在しなかった場合
    '処理終了
    Exit Function
  End If

  ReDim strAryCsvData(intLineCnt - vIntHeaderRowPos, intCommaCount)

  Dim objFS As New FileSystemObject
  Dim objTS As TextStream
  Dim strReadLine As String
  Dim intColCount As Integer
  Dim varAryReadRow As Variant
  Dim intReadRowCnt As Integer
  Dim intReadTargetRowCnt As Integer
  intReadRowCnt = 0
  intReadTargetRowCnt = 0

  Set objTS = objFS.OpenTextFile(strFolderPath)

  Do While objTS.AtEndOfStream <> True
      strReadLine = objTS.ReadLine
    If (vIntHeaderRowPos - 1) <= intReadRowCnt Then
      strReadLine = RemoveCommaInString(strReadLine)
      varAryReadRow = Split(strReadLine, ",")
      
      For intColCount = 0 To UBound(varAryReadRow)
        strAryCsvData(intReadTargetRowCnt, intColCount) = varAryReadRow(intColCount)
      Next
      intReadTargetRowCnt = intReadTargetRowCnt + 1
    End If

    intReadRowCnt = intReadRowCnt + 1
  Loop

  rIntReadColCnt = rIntReadColCnt - 1
  rIntReadRowCnt = intReadTargetRowCnt - 1

  objTS.Close
  Set objFS = Nothing
  Set objTS = Nothing

  ReadCsvFile = strAryCsvData
End Function

'''*************************************************************************************
'''<summary>
''' 文字列の配列から特定文字を削除
'''</summary>
'''<param name="vIntHeaderRowPos">ヘッダー行(何行目か)</param>
'''<param name="vStrEncoding">文字コード</param>
'''<return></return>
'''<option>
'''  Created：2022.04.10 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Sub RemoveSpecificString(ByRef rStrTargetAry() As String, ByVal vStrRemoveStringsLst As VBA.Collection)
  Dim strEditStringTmp As String
  Dim varRemoveStringTmp As Variant
  Dim intRowIndex As Integer
  Dim intColIndex As Integer

  For intRowIndex = 0 To UBound(rStrTargetAry, 1)
    For intColIndex = 0 To UBound(rStrTargetAry, 2)
      strEditStringTmp = rStrTargetAry(intRowIndex, intColIndex)
      For Each varRemoveStringTmp In vStrRemoveStringsLst
        strEditStringTmp = Replace(strEditStringTmp, varRemoveStringTmp, "")
      Next
      rStrTargetAry(intRowIndex, intColIndex) = strEditStringTmp
    Next
  Next

End Sub

'''*************************************************************************************
'''<summary>
''' Excelシートへ二次元配列データを貼付
'''</summary>
'''<param name="vIntHeaderRowPos">ヘッダー行(何行目か)</param>
'''<param name="vStrEncoding">文字コード</param>
'''<return>行数</return>
'''<option>
'''  Created：2022.04.09 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Sub PastAryDataToExcelSheet(ByRef rVarPastEarningsDataAry() As Variant, ByVal vStrTargetSheet As String)
  Dim objSheet As Worksheet
  Dim intEmptyCellNum As Integer
  Dim varPastAmazonEarningsDataAry() As Variant
  Dim intPastMaxColNum As Integer
  Dim intPastMaxRowNum As Integer
  intPastMaxColNum = UBound(rVarPastEarningsDataAry, 2)
  intPastMaxRowNum = UBound(rVarPastEarningsDataAry, 1)
  ReDim varPastAmazonEarningsDataAry(intPastMaxRowNum, intPastMaxColNum)

  intEmptyCellNum = 0
  Set objSheet = Worksheets(vStrTargetSheet)

  Do
    If IsEmpty(objSheet.Cells(intEmptyCellNum + 4, 6).Value) = True Then

      Exit Do
    End If
    intEmptyCellNum = intEmptyCellNum + 1
  Loop

  Dim intRowCnt As Integer
  Dim intRevsRowCnt As Integer
  intRevsRowCnt = intPastMaxRowNum - 1

  '**配列データをソートする処理部(現状並び替えさせていない)**
  For intRowCnt = 0 To intPastMaxRowNum
     varPastAmazonEarningsDataAry(intRowCnt, 0) = rVarPastEarningsDataAry(intRowCnt, 0)
     varPastAmazonEarningsDataAry(intRowCnt, 1) = rVarPastEarningsDataAry(intRowCnt, 1)
     varPastAmazonEarningsDataAry(intRowCnt, 2) = rVarPastEarningsDataAry(intRowCnt, 2)
     varPastAmazonEarningsDataAry(intRowCnt, 3) = rVarPastEarningsDataAry(intRowCnt, 3)
     varPastAmazonEarningsDataAry(intRowCnt, 4) = rVarPastEarningsDataAry(intRowCnt, 4)
     varPastAmazonEarningsDataAry(intRowCnt, 5) = rVarPastEarningsDataAry(intRowCnt, 5)
     varPastAmazonEarningsDataAry(intRowCnt, 6) = rVarPastEarningsDataAry(intRowCnt, 6)
     varPastAmazonEarningsDataAry(intRowCnt, 7) = rVarPastEarningsDataAry(intRowCnt, 7)
     varPastAmazonEarningsDataAry(intRowCnt, 8) = rVarPastEarningsDataAry(intRowCnt, 8)
     varPastAmazonEarningsDataAry(intRowCnt, 9) = rVarPastEarningsDataAry(intRowCnt, 9)
  Next
  '**シート内の貼付始点をCells(x,y)で指定しResizeで貼付範囲を指定し配列をペースト**
  objSheet.Cells(intEmptyCellNum + 4, 6).Resize(intPastMaxRowNum + 1, intPastMaxColNum + 1) = varPastAmazonEarningsDataAry
End Sub

'''*************************************************************************************
'''<summary>
''' ブックオープンチェック処理
'''</summary>
'''<param name=""></param>
'''<return></return>
'''<option>
''' 2021.08.10 K.Hiraoka Created
'''</option>
'''*************************************************************************************

'''*************************************************************************************
'''<summary>
''' コレクションに指定文字列が含まれているかチェック
'''</summary>
'''<param name="vObjCollection">チェック対象のリスト</param>
'''<param name="vStrTargetItem">チェック対象</param>
'''<return>True or False</return>
'''<option>
''' 2021.08.14 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Public Function ContainStrings(ByVal vObjCollection As VBA.Collection, ByVal vStrTargetItem As String) As Boolean
   Dim blnRes As Boolean
   blnRes = False
   
   Dim varItem As Variant
   varItem = ""
   
   '指定文字列がコレクションに含まれているかチェック
   For Each varItem In vObjCollection
      If vStrTargetItem = varItem Then
        blnRes = True
        Exit For
      End If
   Next
   
   ContainStrings = blnRes
End Function

'''*************************************************************************************
'''<summary>
''' コレクションからすべての追加要素を削除
'''</summary>
'''<param name="rObjCollection">削除対象のリスト</param>
'''<return></return>
'''<option>
''' 2021.08.14 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Public Sub RemoveAllLst(ByRef rObjCollection As VBA.Collection)
   Dim varItem As Variant
   varItem = ""
      Dim strItem As String
   strItem = ""
   Dim intItem As Integer
   Dim intItemMax As Integer
   intItemMax = rObjCollection.Count
   
   'セットされているコレクションの要素をすべて削除
      For intItem = intItemMax To 1 Step -1
      rObjCollection.Remove (intItem)
   Next

End Sub

'''*************************************************************************************
'''<summary>
''' ソート処理(バブルソート)
'''</summary>
'''<param name="rObjArgAry">ソート対象の配列</param>
'''<param name="vIntKeyPos">ソートするキーを指定</param>
'''<param name="vIntAscOrDesc">[昇順：0],[降順：1]の指定</param>
'''<return></return>
'''<option>
''' 2021.08.11 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Public Sub BubbleSort(ByRef rObjArgAry As Variant, ByVal vIntKeyPos As Integer, ByVal vIntAscOrDesc As Integer)
   Dim objSwap As Variant
   Dim intRowCnt As Integer
   Dim intColCnt As Integer
   Dim intCnt As Integer
   
   If vIntAscOrDesc = 0 Then
   '昇順に並び替え
    For intRowCnt = LBound(rObjArgAry, 1) To UBound(rObjArgAry, 1)
     For intCnt = LBound(rObjArgAry) To UBound(rObjArgAry) - 1
      If rObjArgAry(intCnt, vIntKeyPos) > rObjArgAry(intCnt + 1, vIntKeyPos) Then
       For intColCnt = LBound(rObjArgAry, 2) To UBound(rObjArgAry, 2)
        objSwap = rObjArgAry(intCnt, intColCnt)
        rObjArgAry(intCnt, intColCnt) = rObjArgAry(intCnt + 1, intColCnt)
        rObjArgAry(intCnt + 1, intColCnt) = objSwap
       Next
      End If
     Next
    Next
   ElseIf vIntAscOrDesc = 1 Then
   '降順に並び替え
    For intRowCnt = LBound(rObjArgAry, 1) To UBound(rObjArgAry, 1)
     For intCnt = LBound(rObjArgAry) To UBound(rObjArgAry) - 1
      If rObjArgAry(intCnt, vIntKeyPos) < rObjArgAry(intCnt + 1, vIntKeyPos) Then
       For intColCnt = LBound(rObjArgAry, 2) To UBound(rObjArgAry, 2)
        objSwap = rObjArgAry(intCnt, intColCnt)
        rObjArgAry(intCnt, intColCnt) = rObjArgAry(intCnt + 1, intColCnt)
        rObjArgAry(intCnt + 1, intColCnt) = objSwap
       Next
      End If
     Next
    Next
   End If
End Sub

'''*************************************************************************************
'''<summary>
''' 指定したファイルの
''' 区切り文字数の取得処理
'''</summary>
'''<param name="vStrFilePath">対象ファイルのフルパス</param>
'''<param name="vStrDelimiter">区切り文字</param>
'''<return>デリミタ数</return>
'''<option>
'''  Created：2022.04.04 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Public Function GetDelimiterCount(ByVal vStrFilePath As String, ByVal vStrDelimiter As String) As Long
    Dim intCount As Integer
    Dim intDelimiterCnt As Integer
    Dim strReadLine As String
    Dim varAry As Variant
    Dim objFS As New FileSystemObject
    Dim objTS As TextStream
    
    '// 引数のファイルが存在しない場合は処理を終了する
    If (objFS.FileExists(vStrFilePath) = False) Then
        intDelimiterCnt = -1
        Exit Function
    End If
      
    'TextStreamオブジェクト作成
    Set objTS = objFS.OpenTextFile(vStrFilePath)
    
    intDelimiterCnt = 0
    Do While objTS.AtEndOfStream <> True
        strReadLine = objTS.ReadLine
        strReadLine = RemoveCommaInString(strReadLine)
        varAry = Split(strReadLine, vStrDelimiter)
        
        If (intDelimiterCnt < UBound(varAry)) Then
            '最大デリミタ数を更新
            intDelimiterCnt = UBound(varAry)
        End If
    Loop
  
    Call objTS.Close
     
    GetDelimiterCount = intDelimiterCnt
End Function


'''*************************************************************************************
'''<summary>
''' 文字列を分解し一文字ずつ配列に格納する
'''</summary>
'''<param name="vStrLineString">分解対象の文字列</param>
'''<return>分解された文字列が格納された配列</return>
'''<option>
'''  Created：2022.04.04 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Public Function DisassembleString(ByVal vStrLineString As String)
  Dim intCnt As Integer
  Dim intLength As Integer
  Dim strCharacterAry() As String

  intLength = Len(vStrLineString)
  ReDim strCharacterAry(intLength - 1) '<<=====配列の要素数として文字数を使用するため[-1]する

  For intCnt = 1 To intLength
    strCharacterAry(intCnt - 1) = Mid(vStrLineString, intCnt, 1)
  Next intCnt

  DisassembleString = strCharacterAry
End Function

'''*************************************************************************************
'''<summary>
''' 文字列から区切り文字以外のカンマを排除
'''</summary>
'''<param name="vStrLineString">対象文字列</param>
'''<return>カンマ排除後の文字列</return>
'''<option>
'''  Created：2022.04.04 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Public Function RemoveCommaInString(ByVal vStrLineString As String) As String
  Dim blnFirstFlag As Boolean
  Dim blnRemoveFlag As Boolean  '<<===区切り文字以外のカンマ排除モードであるか[True:有効, False:無効]
  Dim intCnt As Integer
  Dim intRemoveCnt As Integer
  Dim strCombineString As String
  Dim strAry() As String
  strAry = DisassembleString(vStrLineString)

  blnFirstFlag = True
  blnRemoveFlag = False
  intRemoveCnt = 0
  For intCnt = 0 To UBound(strAry) - 1
    If strAry(intCnt) = Chr(34) Then  '<<===処理中の文字がダブルクォーテーションか判定 ※Chr(34)はダブルクォーテーションを意味
          intRemoveCnt = intRemoveCnt + 1
      If Not blnFirstFlag Then
        If intRemoveCnt > 1 Then
          intRemoveCnt = 1
        End If
      End If
      
      If intRemoveCnt = 1 Then
        If blnRemoveFlag <> True Then
          blnRemoveFlag = True
        Else
          blnRemoveFlag = False
          blnFirstFlag = False
        End If
      ElseIf intRemoveCnt >= 2 Then
          blnRemoveFlag = True
      End If
    Else
      intRemoveCnt = 0  '<<===リセット
    End If

    If blnRemoveFlag Then
      If strAry(intCnt) = "," Then
      Else
        strCombineString = strCombineString & strAry(intCnt)
      End If
    Else
      strCombineString = strCombineString & strAry(intCnt)
    End If
  Next

  RemoveCommaInString = strCombineString
End Function

'''*************************************************************************************
'''<summary>
''' 2次元配列の特定列を抽出し2次元配列を再構成
'''</summary>
'''<param name="vStrBaseAry">加工元の文字列</param>
'''<return>再構成後の2次元配列</return>
'''<option>
'''  Created：2022.04.06 K.Hiraoka
''' Modified：
'''</option>
'''*************************************************************************************
Public Function ReconstructTwoDimArray(ByRef rStrReconstructBaseAry() As String, ByRef rIntExtractColAry() As Integer)
  Dim intColCnt As Integer
  Dim intMaxCol As Integer
  Dim intMaxRow As Integer

  intColCnt = 0
  intMaxCol = UBound(rIntExtractColAry, 1)
  intMaxRow = UBound(rStrReconstructBaseAry, 1)

  Dim strResAry() As String
  ReDim strResAry(intMaxRow, intMaxCol)
  Dim intRowCnt As Integer
  Dim varExtractColInx As Variant
  For intRowCnt = 0 To (intMaxRow)
    For Each varExtractColInx In rIntExtractColAry
        strResAry(intRowCnt, intColCnt) = rStrReconstructBaseAry(intRowCnt, varExtractColInx)
        intColCnt = intColCnt + 1
    Next
    intColCnt = 0
  Next

  ReconstructTwoDimArray = strResAry
End Function
