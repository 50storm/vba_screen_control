Attribute VB_Name = "Message"
Option Explicit


Public Sub showMessage(ctrl As Object, text As String)
     ctrl.Caption = text
     ctrl.BackColor = vbRed
     ctrl.Visible = True
End Sub


Public Sub clearMessage(ctrl As Object)
     ctrl.Caption = ""
     ctrl.Visible = False
End Sub

