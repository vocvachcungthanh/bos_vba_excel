VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formMenu 
   Caption         =   "Cau hinh menu"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16710
   OleObjectBlob   =   "formMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Lay gia tri dong cua danh sach menu
Function getItemMenu()
     With Me
        .txtName.Value = .listMenu_ListBox.List(.listMenu_ListBox.ListIndex, 1)
        .optionParent.Value = .listMenu_ListBox.List(.listMenu_ListBox.ListIndex, 2)
        .txtImage.Value = .listMenu_ListBox.List(.listMenu_ListBox.ListIndex, 3)
        .txtSheet.Value = .listMenu_ListBox.List(.listMenu_ListBox.ListIndex, 4)
    End With
End Function

Private Sub listMenu_ListBox_Change()
  getItemMenu
End Sub


' Xu ly them menu
Function handleCreateMenu()
    Dim name As String
    Dim idParent As String
    Dim image As String
    Dim linkSheet As String
    Dim result As String
    
    If optionParent.ListIndex <> -1 Then
        idParent = optionParent.List(optionParent.ListIndex, 0)
    Else
        idParent = 0
    End If
    
    name = txtName.Value
    image = txtImage.Value
    linkSheet = txtSheet.Value
  
    If name <> "" Then
        result = serverCreateMenu(name, idParent, image, linkSheet)
        serverGetMenu
        Napmenu
        bosMsgBoxSuccess ("Thêm menu")
    Else
        bosMsgBoxError ("Thêm menu")
    End If
End Function

'Xu ly sua menu
Function handleEditMenu(idMenu As String)
     Dim name As String
    Dim idParent As String
    Dim image As String
    Dim linkSheet As String
    Dim result As String
    
    If optionParent.ListIndex <> -1 Then
        idParent = optionParent.List(optionParent.ListIndex, 0)
    Else
        idParent = 0
    End If
    
    name = txtName.Value
    image = txtImage.Value
    linkSheet = txtSheet.Value
  
    If name <> "" Then
        result = serverUpdateMenu(name, idParent, image, linkSheet, idMenu)
        serverGetMenu
        Napmenu
        bosMsgBoxSuccess ("Sua menu")
        
    Else
        bosMsgBoxError ("Sua menu")
    End If
End Function

'Thuc hien luu menu
Private Sub btnSubmit_Click()
    Dim index As String
    Dim id As String
    
    index = listMenu_ListBox.ListIndex
 
    If index <> -1 Then
    id = listMenu_ListBox.List(listMenu_ListBox.ListIndex, 0)
      handleEditMenu (id)
    Else
        handleCreateMenu
    End If
End Sub

'Thuc hien xoa menu
Private Sub btnDelete_Click()
    answ = MsgBox("Ban co trac muong xoa menu nay", vbYesNo)
    
    Debug.Print bgc
    If answ = vbYes Then
        Dim index As String
        Dim id As String
        Dim lr As Long
        lr = ThisWorkbook.Sheets(2).Range("A" & Rows.Count).End(xlUp).row 'Lay dong cuoi cung co du lieu
        
        index = listMenu_ListBox.ListIndex
        
        If index <> -1 Then
             id = listMenu_ListBox.List(listMenu_ListBox.ListIndex, 0)
             serverDeleteMenu (id)
             Napmenu
             bosMsgBoxSuccess ("Xoa menu thanh cong")
             
             For i = 4 To lr
               If Cells(i, 1) = id Then
                    Rows(i).delete
               End If
               
               
             Next i
        End If
    End If
End Sub

Private Sub optionParent_Change()

End Sub

Private Sub UserForm_Click()

End Sub
