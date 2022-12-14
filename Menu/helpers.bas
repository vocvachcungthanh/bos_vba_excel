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
