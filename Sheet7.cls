VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


'///////////////////////////////////////////////////////////
'�w�b�_�[
'///////////////////////////////////////////////////////////

Private Sub chkBox�����֌W�����s�v_Click()
   Call Color.changeColor(chkBox�����֌W�����s�v.Object)
End Sub

Private Sub chkBox�����`�[���ւ̃t�B�[�h�o�b�N_Click()
   Call Color.changeColor(chkBox�����`�[���ւ̃t�B�[�h�o�b�N.Object)
End Sub

'///////////////////////////////////////////////////////////
'Install Information
'///////////////////////////////////////////////////////////
Private Sub chkBox�����֌W�����s�v_Left_Click()
    Call Color.changeColor(chkBox�����֌W�����s�v_Left.Object)
End Sub

Private Sub chkBox��������Ă���_Left_Click()
    Call Color.changeColor(chkBox��������Ă���_Left.Object)
End Sub
Private Sub chkBox�������Ă���_Left_Click()
    Call Color.changeColor(chkBox�������Ă���_Left.Object)
End Sub

Private Sub chkBox��������Ă���_Left_Click()
    Call Color.changeColor(chkBox��������Ă���_Left.Object)
End Sub
Private Sub chkBox�����֌W�����s�v_Right_Click()
    Call Color.changeColor(chkBox�����֌W�����s�v_Right.Object)
End Sub

Private Sub chkBox��������Ă���_Right_Click()
    Call Color.changeColor(chkBox��������Ă���_Right.Object)
End Sub
Private Sub chkBox�������Ă���_Right_Click()
    Call Color.changeColor(chkBox�������Ă���_Right.Object)
End Sub

Private Sub chkBox��������Ă���_Right_Click()
    Call Color.changeColor(chkBox��������Ă���_Right.Object)
End Sub


'///////////////////////////////////////////////////////////
'�������[�̌�����
'///////////////////////////////////////////////////////////
Private Sub chkBox�������Ă悢_Click()
    Call Color.changeColor(chkBox�������Ă悢.Object)
End Sub
Private Sub chkBox�������Ȃ��ł��̂܂ܕԋp_Click()
    Call Color.changeColor(chkBox�������Ȃ��ł��̂܂ܕԋp.Object)
End Sub

'///////////////////////////////////////////////////////////
'�t���[A
'///////////////////////////////////////////////////////////
'(�X�y�A�����v��)
Private Sub chkBox�A���[���{�^��������_Click()
   Call changeColorFor2Ctrl(chkBox�A���[���{�^��������.Object, lbl�A���[���{�^��������.Object)
End Sub

Private Sub chkBox�s�ǂ̂��߃X�y�A�v���\�K�v_Click()
    Call changeColorFor2Ctrl(chkBox�s�ǂ̂��߃X�y�A�v���\�K�v.Object, lbl�s�ǂ̂��߃X�y�A�v���\�K�v.Object)
End Sub

Private Sub chkBox���|�[�g���F_Click()
    Call changeColorFor2Ctrl(chkBox���|�[�g���F.Object, lbl���|�[�g���F.Object)
End Sub

Private Sub chkBox�X�y�A�v���[_Click()
    Call changeColorFor2Ctrl(chkBox�X�y�A�v���[.Object, lbl�X�y�A�v���[.Object)
End Sub

Private Sub chkBox��SN����_Click()
    Call changeColorFor2Ctrl(chkBox��SN����.Object, lbl��SN����.Object)
End Sub
Private Sub chkBoxWEB���SN����_Click()
    Call changeColorFor2Ctrl(chkBoxWEB���SN����.Object, lblWEB���SN����.Object)

End Sub
Private Sub chkBoxTAC�����p�[�c�\��_Click()
    Call changeColorFor2Ctrl(chkBoxTAC�����p�[�c�\��.Object, lblTAC�����p�[�c�\��.Object)
End Sub
Private Sub chkBox�X�y�A�����v���I��_Click()
    Call changeColorFor2Ctrl(chkBox�X�y�A�����v���I��.Object, lbl�X�y�A�����v���I��.Object)
End Sub
Private Sub chkBoxWEB���SN_HOLD����_Click()
    Call changeColorFor2Ctrl(chkBoxWEB���SN_HOLD����.Object, lblWEB���SN_HOLD����.Object)
End Sub
Private Sub chkBoxCisco�l�ɏ��J���m�F_Click()
    Call changeColorFor2Ctrl(chkBoxCisco�l�ɏ��J���m�F.Object, lblCisco�l�ɏ��J���m�F.Object)
End Sub
Private Sub chkBoxSCH��TAC�\������_Click()
    Call changeColorFor2Ctrl(chkBoxSCH��TAC�\������.Object, lblSCH��TAC�\������.Object)
End Sub
Private Sub chkBox����o�׎w��_Click()
    Call changeColorFor2Ctrl(chkBox����o�׎w��.Object, lbl����o�׎w��.Object)
End Sub

Private Sub chkBox�{���y�ь���o�׃��[��_Click()
    Call changeColorFor2Ctrl(chkBox�{���y�ь���o�׃��[��.Object, lbl�{���y�ь���o�׃��[��.Object)
End Sub
Private Sub chkBox�o�׍�Ǝ�_Click()
    Call changeColorFor2Ctrl(chkBox�o�׍�Ǝ�.Object, lbl�o�׍�Ǝ�.Object)
End Sub
Private Sub chkBox���q�l�Ǖi�҂�_Click()
    Call changeColorFor2Ctrl(chkBox���q�l�Ǖi�҂�.Object, lbl���q�l�Ǖi�҂�.Object)
End Sub
Private Sub chkBox�X�y�A�����v����֖����I��_Click()
    Call changeColorFor2Ctrl(chkBox�X�y�A�����v����֖����I��.Object, lbl�X�y�A�����v����֖����I��.Object)
End Sub


'(OSI�i���ARework�s��)
Private Sub chkBokOSI�s�ǂ̂���_Click()
    Call changeColorFor2Ctrl(chkBokOSI�s�ǂ̂���.Object, lblOSI�s�ǂ̂���.Object)
End Sub

Private Sub chkBoxOSI���|�[�g���F_Click()
    Call changeColorFor2Ctrl(chkBoxOSI���|�[�g���F.Object, lblOSI���|�[�g���F.Object)
End Sub

Private Sub chkBoxOSI_notpass�ŏo��_Click()
    Call changeColorFor2Ctrl(chkBoxOSI_notpass�ŏo��.Object, lblOSI_notpass�ŏo��.Object)
End Sub

Private Sub chkBoxOSI_TAC�����p�[�c�\��_Click()
    Call changeColorFor2Ctrl(chkBoxOSI_TAC�����p�[�c�\��.Object, lblOSI_TAC�����p�[�c�\��.Object)
End Sub

Private Sub chkBoxOSI�I��_Click()
    Call changeColorFor2Ctrl(chkBoxOSI�I��.Object, lblOSI�I��.Object)
End Sub


'(Rework)
Private Sub chkBoxRework�X�y�A�v���[_Click()
    Call changeColorFor2Ctrl(chkBoxRework�X�y�A�v���[.Object, lblRework�X�y�A�v���[.Object)
End Sub

Private Sub chkBoxRework�m�F��_Click()
    Call changeColorFor2Ctrl(chkBoxRework�m�F��.Object, lblRework�m�F��.Object)
End Sub

Private Sub chkBoxRework���|�[�g���F_Click()
    Call changeColorFor2Ctrl(chkBoxRework���|�[�g���F.Object, lblRework���|�[�g���F.Object)
End Sub

Private Sub chkBoxRework�I��_Click()
     Call changeColorFor2Ctrl(chkBoxRework�I��.Object, lblRework�I��.Object)
End Sub


'///////////////////////////////////////////////////////////
'���|�[�g�̃`�F�b�N(Upper Side)
'///////////////////////////////////////////////////////////
'Failure Product
Private Sub chkBoxRptChkFailureProduct_Click()
    changeColorForChk (chkBoxRptChkFailureProduct.Object)
End Sub

'Serial Label
Private Sub chkBoxRptChkSerialLabel_Click()
    changeColorForChk (chkBoxRptChkSerialLabel.Object)
End Sub

'Judgement
Private Sub chkBoxRptChkJudgement_Click()
    changeColorForChk (chkBoxRptChkJudgement.Object)
End Sub


'Action
Private Sub chkBoxRptChkAction_Click()
    changeColorForChk (chkBoxRptChkAction.Object)
End Sub


'Defect Category
Private Sub chkBoxRptChkDefectCategory_Click()
    changeColorForChk (chkBoxRptChkDefectCategory.Object)
End Sub


'Defect
Private Sub chkBoxRptChkDefect_Click()
    changeColorForChk (chkBoxRptChkDefect.Object)
End Sub


'Process
Private Sub chkBoxRptChkProcess_Click()
    changeColorForChk (chkBoxRptChkProcess.Object)
End Sub


'�������ꂽPID#
Private Sub chkBoxRptChk�������ꂽPID_Click()
    changeColorForChk (chkBoxRptChk�������ꂽPID.Object)
End Sub

'�������ꂽSN
Private Sub chkBoxRptChk�������ꂽSN_Click()
    changeColorForChk (chkBoxRptChk�������ꂽSN.Object)
End Sub


'///////////////////////////////////////////////////////////
'���|�[�g�̃`�F�b�N(Down Side)
'///////////////////////////////////////////////////////////
Private Sub chkBoxRptChkPackingCondition_Click()
    changeColorForChk (chkBoxRptChkPackingCondition.Object)
End Sub

Private Sub chkBoxRptChk�Č����m�F����_Click()
    changeColorForChk (chkBoxRptChk�Č����m�F����.Object)
End Sub

Private Sub chkBoxRptChk�N���C�e���A_Click()
    changeColorForChk (chkBoxRptChk�N���C�e���A.Object)
End Sub

Private Sub chkBoxRptChk�Ĕ�����_�Ȃ�_Click()
    changeColorForChk (chkBoxRptChk�Ĕ�����_�Ȃ�.Object)
End Sub

Private Sub chkBoxRptChk�ʐ^�G���A�̓��e�p��\��_Click()
    changeColorForChk (chkBoxRptChk�ʐ^�G���A�̓��e�p��\��.Object)
End Sub
Private Sub chkBoxRptChk�f�[�^�x�[�X��Cisco�l�R�����g_Click()
    changeColorForChk (chkBoxRptChk�f�[�^�x�[�X��Cisco�l�R�����g.Object)
End Sub




Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'�K�{�`�F�b�N�����B����Ă����悤�B
'Escort#

'�S�p�𔼊p��
'Range("E7").Value = Trim(Range("E7").Value)
'Range("E7").Value = StrConv(Range("E7").Value, vbNarrow)
'Call Coverter.toHan(Range("E7"))

If Trim(Range("E6").Value) = "" Then
    Call changeColorForRangeFalse(Range("E6"))
    Call Message.showMessage(lblMsg, "[�K�{����]" & vbCrLf & "Escort#�����͂���Ă��܂���!!!!")
    Range("E6").Select
    Exit Sub
Else
    Call changeColorForRangeOk(Range("E6"))
    Call clearMessage(lblMsg)
    
End If

'SO/SS
If Trim(Range("AH6").Value) = "" Then
    Call changeColorForRangeFalse(Range("AH6"))
    Call Message.showMessage(lblMsg, "[�K�{����]" & vbCrLf & "SO/SS�����͂���Ă��܂���!!!!")
'    MsgBox ("SO/SS�����͂���Ă��܂���B")
    Range("AH6").Select
    Exit Sub
    
Else
    Call changeColorForRangeOk(Range("AH6"))
     Call clearMessage(lblMsg)
End If

'Failure PID
If Trim(Range("E11").Value) = "" Then
    Call changeColorForRangeFalse(Range("E11"))
'    MsgBox ("Failure PID#�����͂���Ă��܂���B")
    Call Message.showMessage(lblMsg, "[�K�{����]" & vbCrLf & "Failure PID#�����͂���Ă��܂���!!!!")
    Range("E11").Select
    Exit Sub
Else
    Call changeColorForRangeOk(Range("E11"))
    Call clearMessage(lblMsg)

End If

'���iSerial#
If Trim(Range("AH11").Value) = "" Then
    Call changeColorForRangeFalse(Range("AH11"))
'    MsgBox ("���iSerial#�����͂���Ă��܂���B")
    Call Message.showMessage(lblMsg, "[�K�{����]" & vbCrLf & "���iSerial#�����͂���Ă��܂���!!!!")
    Range("AH11").Select
    Exit Sub

Else
    Call changeColorForRangeOk(Range("AH11"))
    Call clearMessage(lblMsg)

End If

'�S�p�𔼊p��
Range("E6").Value = Trim(Range("E6").Value)
Range("E6").Value = StrConv(Range("E6").Value, vbNarrow + vbUpperCase)
Range("AH6").Value = Trim(Range("AH6").Value)
Range("AH6").Value = StrConv(Range("AH6").Value, vbNarrow + vbUpperCase)
Range("E11").Value = Trim(Range("E11").Value)
Range("E11").Value = StrConv(Range("E11").Value, vbNarrow + vbUpperCase)
Range("AH11").Value = Trim(Range("AH11").Value)
Range("AH11").Value = StrConv(Range("AH11").Value, vbNarrow + vbUpperCase)




End Sub
