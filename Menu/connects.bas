Attribute VB_Name = "connects"
Function Connect(sqlQuery As String)
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim rs As ADODB.Recordset
    Dim SQLStr As String
    
    Set Cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
   
    
    StrCnn = "Driver={SQL Server};Server=192.168.100.72,1433;Database=bos_vba;Uid=bosdev;Pwd=bos123456;"
    'Xu ly lenh
    SQLStr = sqlQuery
  
    Cn.Open StrCnn
    rs.Open SQLStr, Cn, adOpenStatic
        
    Cn.Close
    Set Cn = Nothing
End Function
