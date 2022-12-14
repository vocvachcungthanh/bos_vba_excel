Attribute VB_Name = "connects"
Function Connect(sqlQuery As String)
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim rs As ADODB.RecordSet
    Dim SQLStr As String
    
    Set Cn = New ADODB.Connection
    Set rs = New ADODB.RecordSet
   
    
    StrCnn = "Driver={SQL Server};Server=192.168.100.72,1433;Database=bos_vba;Uid=bosdev;Pwd=bos123456;"
  
    'Xu ly lenh
    SQLStr = sqlQuery
  
    Cn.Open StrCnn
    rs.Open SQLStr, Cn, adOpenStatic
        
    Cn.Close
    Set Cn = Nothing
End Function


Function sqlGetRows(sqlQuery As String) As Variant
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim rs As ADODB.RecordSet
    Set rs = New ADODB.RecordSet
 
    StrCnn = "Driver={SQL Server};Server=192.168.100.72,1433;Database=bos_vba;Uid=bosdev;Pwd=bos123456;"
    Debug.Print sqlQuery
    'Xu ly lenh
  
    SQLStr = sqlQuery
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    rs.Open SQLStr, Cn, adOpenStatic
       sqlGetRows = rs.GetRows()
    Cn.Close
    Set Cn = Nothing
End Function
