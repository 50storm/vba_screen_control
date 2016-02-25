Attribute VB_Name = "Color"
Option Explicit
'呼ぶ側 Callを付けないと、引数複数や、オブジェクトを渡せない。イケてないいＶＢＡの仕様。。。。
Public Sub changeColor(ctrl As Object)
     If ctrl.Value Then
         ctrl.BackColor = vbYellow
     Else
         ctrl.BackColor = vbWhite
     End If
End Sub

Public Sub changeColorForChk(ctrl As Object)
     If ctrl.Value Then
        ctrl.BackColor = RGB(255, 153, 205)

     Else
        ctrl.BackColor = RGB(255, 255, 153)
              
     End If
End Sub


Public Sub changeColorFor2Ctrl(chkCtrl As Object, targetCtrl As Object)

     If chkCtrl.Value Then
         targetCtrl.BackColor = vbYellow
     Else
         targetCtrl.BackColor = vbWhite
     End If

End Sub

Public Sub changeColorForRangeFalse(rng As Range)
    rng.Interior.Color = vbRed
End Sub


Public Sub changeColorForRangeOk(rng As Range)
    rng.Interior.Color = RGB(255, 255, 153)
End Sub

