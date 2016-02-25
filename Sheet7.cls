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
'ヘッダー
'///////////////////////////////////////////////////////////

Private Sub chkBox実装関係無し不要_Click()
   Call Color.changeColor(chkBox実装関係無し不要.Object)
End Sub

Private Sub chkBox検査チームへのフィードバック_Click()
   Call Color.changeColor(chkBox検査チームへのフィードバック.Object)
End Sub

'///////////////////////////////////////////////////////////
'Install Information
'///////////////////////////////////////////////////////////
Private Sub chkBox実装関係無し不要_Left_Click()
    Call Color.changeColor(chkBox実装関係無し不要_Left.Object)
End Sub

Private Sub chkBox実装されていた_Left_Click()
    Call Color.changeColor(chkBox実装されていた_Left.Object)
End Sub
Private Sub chkBox実装していた_Left_Click()
    Call Color.changeColor(chkBox実装していた_Left.Object)
End Sub

Private Sub chkBox同封されていた_Left_Click()
    Call Color.changeColor(chkBox同封されていた_Left.Object)
End Sub
Private Sub chkBox実装関係無し不要_Right_Click()
    Call Color.changeColor(chkBox実装関係無し不要_Right.Object)
End Sub

Private Sub chkBox実装されていた_Right_Click()
    Call Color.changeColor(chkBox実装されていた_Right.Object)
End Sub
Private Sub chkBox実装していた_Right_Click()
    Call Color.changeColor(chkBox実装していた_Right.Object)
End Sub

Private Sub chkBox同封されていた_Right_Click()
    Call Color.changeColor(chkBox同封されていた_Right.Object)
End Sub


'///////////////////////////////////////////////////////////
'メモリーの交換可否
'///////////////////////////////////////////////////////////
Private Sub chkBox交換してよい_Click()
    Call Color.changeColor(chkBox交換してよい.Object)
End Sub
Private Sub chkBox交換しないでそのまま返却_Click()
    Call Color.changeColor(chkBox交換しないでそのまま返却.Object)
End Sub

'///////////////////////////////////////////////////////////
'フローA
'///////////////////////////////////////////////////////////
'(スペア交換要求)
Private Sub chkBoxアラームボタンを押す_Click()
   Call changeColorFor2Ctrl(chkBoxアラームボタンを押す.Object, lblアラームボタンを押す.Object)
End Sub

Private Sub chkBox不良のためスペア要求表必要_Click()
    Call changeColorFor2Ctrl(chkBox不良のためスペア要求表必要.Object, lbl不良のためスペア要求表必要.Object)
End Sub

Private Sub chkBoxレポート承認_Click()
    Call changeColorFor2Ctrl(chkBoxレポート承認.Object, lblレポート承認.Object)
End Sub

Private Sub chkBoxスペア要求票_Click()
    Call changeColorFor2Ctrl(chkBoxスペア要求票.Object, lblスペア要求票.Object)
End Sub

Private Sub chkBox箱SN交換_Click()
    Call changeColorFor2Ctrl(chkBox箱SN交換.Object, lbl箱SN交換.Object)
End Sub
Private Sub chkBoxWEB代替SN入力_Click()
    Call changeColorFor2Ctrl(chkBoxWEB代替SN入力.Object, lblWEB代替SN入力.Object)

End Sub
Private Sub chkBoxTAC交換パーツ申請_Click()
    Call changeColorFor2Ctrl(chkBoxTAC交換パーツ申請.Object, lblTAC交換パーツ申請.Object)
End Sub
Private Sub chkBoxスペア交換要求終了_Click()
    Call changeColorFor2Ctrl(chkBoxスペア交換要求終了.Object, lblスペア交換要求終了.Object)
End Sub
Private Sub chkBoxWEB代替SN_HOLD入力_Click()
    Call changeColorFor2Ctrl(chkBoxWEB代替SN_HOLD入力.Object, lblWEB代替SN_HOLD入力.Object)
End Sub
Private Sub chkBoxCisco様に情報開示確認_Click()
    Call changeColorFor2Ctrl(chkBoxCisco様に情報開示確認.Object, lblCisco様に情報開示確認.Object)
End Sub
Private Sub chkBoxSCHへTAC申請情報提供_Click()
    Call changeColorFor2Ctrl(chkBoxSCHへTAC申請情報提供.Object, lblSCHへTAC申請情報提供.Object)
End Sub
Private Sub chkBox現状出荷指示_Click()
    Call changeColorFor2Ctrl(chkBox現状出荷指示.Object, lbl現状出荷指示.Object)
End Sub

Private Sub chkBox本紙及び現状出荷メール_Click()
    Call changeColorFor2Ctrl(chkBox本紙及び現状出荷メール.Object, lbl本紙及び現状出荷メール.Object)
End Sub
Private Sub chkBox出荷作業者_Click()
    Call changeColorFor2Ctrl(chkBox出荷作業者.Object, lbl出荷作業者.Object)
End Sub
Private Sub chkBoxお客様良品待ち_Click()
    Call changeColorFor2Ctrl(chkBoxお客様良品待ち.Object, lblお客様良品待ち.Object)
End Sub
Private Sub chkBoxスペア交換要求代替無し終了_Click()
    Call changeColorFor2Ctrl(chkBoxスペア交換要求代替無し終了.Object, lblスペア交換要求代替無し終了.Object)
End Sub


'(OSI品かつ、Rework不可)
Private Sub chkBokOSI不良のため_Click()
    Call changeColorFor2Ctrl(chkBokOSI不良のため.Object, lblOSI不良のため.Object)
End Sub

Private Sub chkBoxOSIレポート承認_Click()
    Call changeColorFor2Ctrl(chkBoxOSIレポート承認.Object, lblOSIレポート承認.Object)
End Sub

Private Sub chkBoxOSI_notpassで出荷_Click()
    Call changeColorFor2Ctrl(chkBoxOSI_notpassで出荷.Object, lblOSI_notpassで出荷.Object)
End Sub

Private Sub chkBoxOSI_TAC交換パーツ申請_Click()
    Call changeColorFor2Ctrl(chkBoxOSI_TAC交換パーツ申請.Object, lblOSI_TAC交換パーツ申請.Object)
End Sub

Private Sub chkBoxOSI終了_Click()
    Call changeColorFor2Ctrl(chkBoxOSI終了.Object, lblOSI終了.Object)
End Sub


'(Rework)
Private Sub chkBoxReworkスペア要求票_Click()
    Call changeColorFor2Ctrl(chkBoxReworkスペア要求票.Object, lblReworkスペア要求票.Object)
End Sub

Private Sub chkBoxRework確認者_Click()
    Call changeColorFor2Ctrl(chkBoxRework確認者.Object, lblRework確認者.Object)
End Sub

Private Sub chkBoxReworkレポート承認_Click()
    Call changeColorFor2Ctrl(chkBoxReworkレポート承認.Object, lblReworkレポート承認.Object)
End Sub

Private Sub chkBoxRework終了_Click()
     Call changeColorFor2Ctrl(chkBoxRework終了.Object, lblRework終了.Object)
End Sub


'///////////////////////////////////////////////////////////
'レポートのチェック(Upper Side)
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


'実装されたPID#
Private Sub chkBoxRptChk実装されたPID_Click()
    changeColorForChk (chkBoxRptChk実装されたPID.Object)
End Sub

'実装されたSN
Private Sub chkBoxRptChk実装されたSN_Click()
    changeColorForChk (chkBoxRptChk実装されたSN.Object)
End Sub


'///////////////////////////////////////////////////////////
'レポートのチェック(Down Side)
'///////////////////////////////////////////////////////////
Private Sub chkBoxRptChkPackingCondition_Click()
    changeColorForChk (chkBoxRptChkPackingCondition.Object)
End Sub

Private Sub chkBoxRptChk再現性確認結果_Click()
    changeColorForChk (chkBoxRptChk再現性確認結果.Object)
End Sub

Private Sub chkBoxRptChkクライテリア_Click()
    changeColorForChk (chkBoxRptChkクライテリア.Object)
End Sub

Private Sub chkBoxRptChk再発あり_なし_Click()
    changeColorForChk (chkBoxRptChk再発あり_なし.Object)
End Sub

Private Sub chkBoxRptChk写真エリアの内容英語表現_Click()
    changeColorForChk (chkBoxRptChk写真エリアの内容英語表現.Object)
End Sub
Private Sub chkBoxRptChkデータベース側Cisco様コメント_Click()
    changeColorForChk (chkBoxRptChkデータベース側Cisco様コメント.Object)
End Sub




Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'必須チェックだけ。入れてあげよう。
'Escort#

'全角を半角に
'Range("E7").Value = Trim(Range("E7").Value)
'Range("E7").Value = StrConv(Range("E7").Value, vbNarrow)
'Call Coverter.toHan(Range("E7"))

If Trim(Range("E6").Value) = "" Then
    Call changeColorForRangeFalse(Range("E6"))
    Call Message.showMessage(lblMsg, "[必須入力]" & vbCrLf & "Escort#が入力されていません!!!!")
    Range("E6").Select
    Exit Sub
Else
    Call changeColorForRangeOk(Range("E6"))
    Call clearMessage(lblMsg)
    
End If

'SO/SS
If Trim(Range("AH6").Value) = "" Then
    Call changeColorForRangeFalse(Range("AH6"))
    Call Message.showMessage(lblMsg, "[必須入力]" & vbCrLf & "SO/SSが入力されていません!!!!")
'    MsgBox ("SO/SSが入力されていません。")
    Range("AH6").Select
    Exit Sub
    
Else
    Call changeColorForRangeOk(Range("AH6"))
     Call clearMessage(lblMsg)
End If

'Failure PID
If Trim(Range("E11").Value) = "" Then
    Call changeColorForRangeFalse(Range("E11"))
'    MsgBox ("Failure PID#が入力されていません。")
    Call Message.showMessage(lblMsg, "[必須入力]" & vbCrLf & "Failure PID#が入力されていません!!!!")
    Range("E11").Select
    Exit Sub
Else
    Call changeColorForRangeOk(Range("E11"))
    Call clearMessage(lblMsg)

End If

'製品Serial#
If Trim(Range("AH11").Value) = "" Then
    Call changeColorForRangeFalse(Range("AH11"))
'    MsgBox ("製品Serial#が入力されていません。")
    Call Message.showMessage(lblMsg, "[必須入力]" & vbCrLf & "製品Serial#が入力されていません!!!!")
    Range("AH11").Select
    Exit Sub

Else
    Call changeColorForRangeOk(Range("AH11"))
    Call clearMessage(lblMsg)

End If

'全角を半角に
Range("E6").Value = Trim(Range("E6").Value)
Range("E6").Value = StrConv(Range("E6").Value, vbNarrow + vbUpperCase)
Range("AH6").Value = Trim(Range("AH6").Value)
Range("AH6").Value = StrConv(Range("AH6").Value, vbNarrow + vbUpperCase)
Range("E11").Value = Trim(Range("E11").Value)
Range("E11").Value = StrConv(Range("E11").Value, vbNarrow + vbUpperCase)
Range("AH11").Value = Trim(Range("AH11").Value)
Range("AH11").Value = StrConv(Range("AH11").Value, vbNarrow + vbUpperCase)




End Sub

