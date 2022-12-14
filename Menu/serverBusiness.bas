Attribute VB_Name = "serverBusiness"
'Lay danh sach ke hoach danh thu
Function getBusinessPlan()
  Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim result As Long
    
    StrCnn = "Driver={SQL Server};Server=192.168.100.72,1433;Database=bos_vba;Uid=bosdev;Pwd=bos123456;"
 
    'Xu ly lenh
    Dim SQLStr As String
    SQLStr = "SELECT * FROM Ds_KheHoachKinhDoanhDonVi"
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    rs.Open SQLStr, Cn, adOpenStatic
        ThisWorkbook.Sheets(16).Range("A2").CopyFromRecordset rs
    Cn.Close
    Set Cn = Nothing
    
    
End Function
