Attribute VB_Name = "serveDepartment"
Function getListDepartment() As Variant


    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    StrCnn = "Driver={SQL Server};Server=192.168.100.72,1433;Database=bos_vba;Uid=bosdev;Pwd=bos123456;"
    
    'Xu ly lenh
    Dim SQLStr As String
    SQLStr = "SELECT * FROM PhongBan"
  
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    rs.Open SQLStr, Cn, adOpenStatic
        getListDepartment = rs.GetRows
    Cn.Close
    Set Cn = Nothing
End Function


Function getListDepartmentParent(id As Integer) As Variant
   
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    StrCnn = "Driver={SQL Server};Server=192.168.100.72,1433;Database=bos_vba;Uid=bosdev;Pwd=bos123456;"
    
    'Xu ly lenh
    Dim SQLStr As String
    SQLStr = "SELECT * FROM PhongBan WHERE CapPhongBan =" & id & ""
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    rs.Open SQLStr, Cn, adOpenStatic
       getListDepartmentParent = rs.GetRows
    Cn.Close
    Set Cn = Nothing
End Function
