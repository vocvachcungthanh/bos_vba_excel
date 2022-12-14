Attribute VB_Name = "helpers"
Function bosMsgBoxSuccess(title As String)
    MsgBox "thanh cong", , title
End Function

Function bosMsgBoxError(title As String)
    MsgBox "That bai", , title
End Function

Public Function GetArrLength(a As Variant) As Long
   If IsEmpty(a) Then
      GetArrLength = 0
   Else
      GetArrLength = UBound(a) - LBound(a) + 1
   End If
End Function

Sub ThongBao(noidung As String)
    Dim tieude As String
    
    tieude = "Thông báo"
    Application.Assistant.DoAlert tieude, noidung, msoAlertButtonOK, msoAlertIconInfo, msoAlertDefaultFirst, 0, False
    
End Sub
