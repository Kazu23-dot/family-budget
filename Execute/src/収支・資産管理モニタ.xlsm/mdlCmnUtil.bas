Attribute VB_Name = "mdlCmnUtil"
Option Explicit

'''*************************************************************************************
'''<summary>
''' �A���������R�[�h�����擾
''' ���[�N�V�[�g,�J�n��,�J�n�s���w�肵�������s
'''</summary>
'''<param name="vObjWorkSheet">�Ώۃ��[�N�V�[�g��</param>
'''<param name="vStrStartCol">�J�n��</param>
'''<param name="vIntStartRow">�J�n�s</param>
'''<return>���R�[�h��</return>
'''<option>
''' 2021.06.12 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Public Function GetRowCnt(vObjWorkSheet As Worksheet, vStrStartCol As String, vIntStartRow As Integer) As Integer
 Dim blnContinue As Boolean
 Dim intRowCnt As Integer
 
 '�����l�ݒ�
 blnContinue = True
 intRowCnt = 0
 
 '��f�[�^�𔭌�����܂Ŗ������[�v
 Do While blnContinue
   If vObjWorkSheet.Range(vStrStartCol & (vIntStartRow + intRowCnt)).Value <> "" Then
     intRowCnt = intRowCnt + 1
   Else
     '�J�E���g�I��
     blnContinue = False
   End If
 Loop
   
 '�߂�l��ݒ�
 GetRowCnt = intRowCnt
End Function

'''*************************************************************************************
'''<summary>
''' �w��t�@�C������s�����擾
'''</summary>
'''<param name="vStrFilePath"></param>
'''<return>�s��</return>
'''<option>
''' 2021.06.12 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Public Function GetFileLineCount(vStrFilePath As String) As Long
    Dim objFS As New FileSystemObject
    Dim objTS As TextStream
    
    '�����̃t�@�C�������݂��Ȃ��ꍇ�͏������I��
    If (objFS.FileExists(vStrFilePath) = False) Then
        GetFileLineCount = -1
        Exit Function
    End If
    
    '�ǉ����[�h�ŊJ��
    Set objTS = objFS.OpenTextFile(vStrFilePath, ForAppending)
    
    '�߂�l��ݒ�
    GetFileLineCount = objTS.Line - 1
End Function

'''*************************************************************************************
'''<summary>
''' �w�肳�ꂽCSV����f�[�^�擾
'''</summary>
'''<param name="vIntHeaderRowPos">�w�b�_�[�s(���s�ڂ�)</param>
'''<param name="vStrEncoding">�����R�[�h</param>
'''<return>�s��</return>
'''<option>
'''  Created�F2022.04.03 K.Hiraoka
''' Modified�F
'''</option>
'''*************************************************************************************
Public Function ReadCsvFile(ByVal vIntHeaderRowPos As Integer, ByVal vStrEncoding As String, ByVal vStrDelimiter As String, _
                            ByRef rIntReadColCnt As Integer, ByRef rIntReadRowCnt As Integer)
  Dim blnRes As String
  Dim strAryCsvData() As String

  Dim strFolderPath As String
  Dim objMsgRslt As VbMsgBoxResult
  objMsgRslt = MsgBox(("CSV�t�@�C���̃C���|�[�g�����s���܂����H"), vbYesNo + vbQuestion, "�C���|�[�g�m�F")

  '�C���|�[�g�𒆎~����ꍇ
  If objMsgRslt = vbNo Then
     '�����I��
     Exit Function
  End If

  If Application.FileDialog(msoFileDialogFilePicker).Show = -1 Then
      '�t�@�C���I�����OK���������ꂽ�ꍇ
      strFolderPath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
  Else
      '�L�����Z�����I�����ꂽ�ꍇ
      MsgBox "�����𒆎~���܂�", vbCritical
      Exit Function
  End If

  Dim intLineCnt As Integer
  intLineCnt = GetFileLineCount(strFolderPath)
  
  Dim intCommaCount As Integer
  intCommaCount = GetDelimiterCount(strFolderPath, vStrDelimiter)

  If (intLineCnt = -1) Or (intCommaCount = -1) Then
    '�t�@�C���Ƀf�[�^�����݂��Ȃ������ꍇ
    '�����I��
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
''' ������̔z�񂩂���蕶�����폜
'''</summary>
'''<param name="vIntHeaderRowPos">�w�b�_�[�s(���s�ڂ�)</param>
'''<param name="vStrEncoding">�����R�[�h</param>
'''<return></return>
'''<option>
'''  Created�F2022.04.10 K.Hiraoka
''' Modified�F
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
''' Excel�V�[�g�֓񎟌��z��f�[�^��\�t
'''</summary>
'''<param name="vIntHeaderRowPos">�w�b�_�[�s(���s�ڂ�)</param>
'''<param name="vStrEncoding">�����R�[�h</param>
'''<return>�s��</return>
'''<option>
'''  Created�F2022.04.09 K.Hiraoka
''' Modified�F
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

  '**�z��f�[�^���\�[�g���鏈����(������ёւ������Ă��Ȃ�)**
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
  '**�V�[�g���̓\�t�n�_��Cells(x,y)�Ŏw�肵Resize�œ\�t�͈͂��w�肵�z����y�[�X�g**
  objSheet.Cells(intEmptyCellNum + 4, 6).Resize(intPastMaxRowNum + 1, intPastMaxColNum + 1) = varPastAmazonEarningsDataAry
End Sub

'''*************************************************************************************
'''<summary>
''' �u�b�N�I�[�v���`�F�b�N����
'''</summary>
'''<param name=""></param>
'''<return></return>
'''<option>
''' 2021.08.10 K.Hiraoka Created
'''</option>
'''*************************************************************************************

'''*************************************************************************************
'''<summary>
''' �R���N�V�����Ɏw�蕶���񂪊܂܂�Ă��邩�`�F�b�N
'''</summary>
'''<param name="vObjCollection">�`�F�b�N�Ώۂ̃��X�g</param>
'''<param name="vStrTargetItem">�`�F�b�N�Ώ�</param>
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
   
   '�w�蕶���񂪃R���N�V�����Ɋ܂܂�Ă��邩�`�F�b�N
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
''' �R���N�V�������炷�ׂĂ̒ǉ��v�f���폜
'''</summary>
'''<param name="rObjCollection">�폜�Ώۂ̃��X�g</param>
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
   
   '�Z�b�g����Ă���R���N�V�����̗v�f�����ׂč폜
      For intItem = intItemMax To 1 Step -1
      rObjCollection.Remove (intItem)
   Next

End Sub

'''*************************************************************************************
'''<summary>
''' �\�[�g����(�o�u���\�[�g)
'''</summary>
'''<param name="rObjArgAry">�\�[�g�Ώۂ̔z��</param>
'''<param name="vIntKeyPos">�\�[�g����L�[���w��</param>
'''<param name="vIntAscOrDesc">[�����F0],[�~���F1]�̎w��</param>
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
   '�����ɕ��ёւ�
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
   '�~���ɕ��ёւ�
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
''' �w�肵���t�@�C����
''' ��؂蕶�����̎擾����
'''</summary>
'''<param name="vStrFilePath">�Ώۃt�@�C���̃t���p�X</param>
'''<param name="vStrDelimiter">��؂蕶��</param>
'''<return>�f���~�^��</return>
'''<option>
'''  Created�F2022.04.04 K.Hiraoka
''' Modified�F
'''</option>
'''*************************************************************************************
Public Function GetDelimiterCount(ByVal vStrFilePath As String, ByVal vStrDelimiter As String) As Long
    Dim intCount As Integer
    Dim intDelimiterCnt As Integer
    Dim strReadLine As String
    Dim varAry As Variant
    Dim objFS As New FileSystemObject
    Dim objTS As TextStream
    
    '// �����̃t�@�C�������݂��Ȃ��ꍇ�͏������I������
    If (objFS.FileExists(vStrFilePath) = False) Then
        intDelimiterCnt = -1
        Exit Function
    End If
      
    'TextStream�I�u�W�F�N�g�쐬
    Set objTS = objFS.OpenTextFile(vStrFilePath)
    
    intDelimiterCnt = 0
    Do While objTS.AtEndOfStream <> True
        strReadLine = objTS.ReadLine
        strReadLine = RemoveCommaInString(strReadLine)
        varAry = Split(strReadLine, vStrDelimiter)
        
        If (intDelimiterCnt < UBound(varAry)) Then
            '�ő�f���~�^�����X�V
            intDelimiterCnt = UBound(varAry)
        End If
    Loop
  
    Call objTS.Close
     
    GetDelimiterCount = intDelimiterCnt
End Function


'''*************************************************************************************
'''<summary>
''' ������𕪉����ꕶ�����z��Ɋi�[����
'''</summary>
'''<param name="vStrLineString">����Ώۂ̕�����</param>
'''<return>�������ꂽ�����񂪊i�[���ꂽ�z��</return>
'''<option>
'''  Created�F2022.04.04 K.Hiraoka
''' Modified�F
'''</option>
'''*************************************************************************************
Public Function DisassembleString(ByVal vStrLineString As String)
  Dim intCnt As Integer
  Dim intLength As Integer
  Dim strCharacterAry() As String

  intLength = Len(vStrLineString)
  ReDim strCharacterAry(intLength - 1) '<<=====�z��̗v�f���Ƃ��ĕ��������g�p���邽��[-1]����

  For intCnt = 1 To intLength
    strCharacterAry(intCnt - 1) = Mid(vStrLineString, intCnt, 1)
  Next intCnt

  DisassembleString = strCharacterAry
End Function

'''*************************************************************************************
'''<summary>
''' �����񂩂��؂蕶���ȊO�̃J���}��r��
'''</summary>
'''<param name="vStrLineString">�Ώە�����</param>
'''<return>�J���}�r����̕�����</return>
'''<option>
'''  Created�F2022.04.04 K.Hiraoka
''' Modified�F
'''</option>
'''*************************************************************************************
Public Function RemoveCommaInString(ByVal vStrLineString As String) As String
  Dim blnFirstFlag As Boolean
  Dim blnRemoveFlag As Boolean  '<<===��؂蕶���ȊO�̃J���}�r�����[�h�ł��邩[True:�L��, False:����]
  Dim intCnt As Integer
  Dim intRemoveCnt As Integer
  Dim strCombineString As String
  Dim strAry() As String
  strAry = DisassembleString(vStrLineString)

  blnFirstFlag = True
  blnRemoveFlag = False
  intRemoveCnt = 0
  For intCnt = 0 To UBound(strAry) - 1
    If strAry(intCnt) = Chr(34) Then  '<<===�������̕������_�u���N�H�[�e�[�V���������� ��Chr(34)�̓_�u���N�H�[�e�[�V�������Ӗ�
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
      intRemoveCnt = 0  '<<===���Z�b�g
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
''' 2�����z��̓����𒊏o��2�����z����č\��
'''</summary>
'''<param name="vStrBaseAry">���H���̕�����</param>
'''<return>�č\�����2�����z��</return>
'''<option>
'''  Created�F2022.04.06 K.Hiraoka
''' Modified�F
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
