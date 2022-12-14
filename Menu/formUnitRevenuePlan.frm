VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formUnitRevenuePlan 
   Caption         =   "Ke Hoach Danh Thu Don VI"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17850
   OleObjectBlob   =   "formUnitRevenuePlan.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "formUnitRevenuePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub selectPhongBan_Enter()
 Dim rs As Variant
    Dim i As Integer
    rs = getListDepartment
    
    For i = LBound(rs) To UBound(rs)
        If GetArrLength(rs) > 0 Then
            'selectPhongBan.AddItem rs(0, i)
        End If
        
    Next i
End Sub
