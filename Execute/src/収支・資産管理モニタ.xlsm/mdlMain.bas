Attribute VB_Name = "mdlMain"

'''*************************************************************************************
'''<summary>
''' ����,�x�o,���Y,���f�[�^�����W
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
''' ���ԃR���{�̑I�����𐶐� �����ݔN����ߋ�40�N���I�����𐶐�
'''</summary>
'''<return></return>
'''<option>
''' 2022.05.05 K.Hiraoka Created
'''</option>
'''*************************************************************************************
Sub CreateYearCmbSelection()
 '<<=============================================================>>
 '<<==[1]�Ώۃe�[�u�����̃f�[�^���N���A============================>>
 '<<=============================================================>>
    Dim objDeleteTable As ListObject
    Set objDeleteTable = ThisWorkbook.Sheets("Settings").ListObjects("SelectYearCmb")
    If Not objDeleteTable.DataBodyRange Is Nothing Then
        objDeleteTable.DataBodyRange.Delete
    End If

 '<<=============================================================>>
 '<<==[2]���ݓ�������N�����擾���ߋ�40�N���̑I�����𐶐�==========>>
 '<<=============================================================>>
    Dim intYearTmp As Integer
    Dim strNowYear As String
    strNowYear = CStr(Year(Now))
    intYearTmp = strNowYear

    With ThisWorkbook.Sheets("Settings").ListObjects("SelectYearCmb")
        For intRowCnt = 1 To 40
            .ListRows.Add (1)
            .ListColumns("�I����").DataBodyRange(1) = CStr(intYearTmp)
            intYearTmp = CInt(strNowYear) - intRowCnt
        Next
    End With
 '<<=============================================================>>
 '<<==[3]���������N�R���{�{�b�N�X�p�̃e�[�u�����\�[�g================>>
 '<<=============================================================>>
    With ThisWorkbook.Sheets("Settings").ListObjects("SelectYearCmb")
        .Range.Sort key1:=.ListColumns("�I����").Range, order1:=xlDescending, Header:=xlYes
    End With
End Sub

'''*************************************************************************************
'''<summary>
''' �e�X�g���s�p�̃��\�b�h
'''</summary>
'''<param name="vIntHeaderRowPos">�w�b�_�[�s(���s�ڂ�)</param>
'''<param name="vStrEncoding">�����R�[�h</param>
'''<return>�s��</return>
'''<option>
'''  Created�F2022.04.03 K.Hiraoka
''' Modified�F
'''</option>
'''*************************************************************************************
Public Sub ExecuteTestMode()
    With ThisWorkbook.Sheets("DataSource").ListObjects("DataSource")
        .Range.Sort key1:=.ListColumns("���C���J�e�S��").Range, order1:=xlDescending, _
                    key2:=.ListColumns("���x�z�E���Y���z").Range, order2:=xlDescending, Header:=xlYes
    End With
End Sub
