Attribute VB_Name = "serverMenu"
Function serverGetMenu()
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    StrCnn = "Driver={SQL Server};Server=192.168.100.72,1433;Database=bos_vba;Uid=bosdev;Pwd=bos123456;"
    
    'Xu ly lenh
    Dim SQLStr As String
    SQLStr = "SELECT * FROM menu"
  
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    rs.Open SQLStr, Cn, adOpenStatic
        ThisWorkbook.Sheets(2).Range("A4").CopyFromRecordset rs
    Cn.Close
    Set Cn = Nothing
End Function

Function serverCreateMenu(name As String, idParent As String, image As String, linkSheet As String)
    Dim SQLStr As String
    Dim Kt As String 'Khai bao ky tu nhay don "'"
    Kt = "'"
    
    'Xu ly lenh
    SQLStr = "INSERT INTO menu(name,id_parent,image, link_sheet) VALUES(N" & Kt & name & Kt & "," & Kt & idParent & Kt & "," & Kt & image & Kt & ", N" & Kt & linkSheet & Kt & ")"
    Connect (SQLStr)
End Function

Function serverUpdateMenu(name As String, idParent As String, image As String, linkSheet As String, idMenu As String)
    Dim SQLStr As String
    Dim Kt As String 'Khai bao ky tu nhay don "'"
    Kt = "'"
    
    SQLStr = "UPDATE menu SET name= N" & Kt & name & Kt & "," & "id_parent" & "=" & Kt & idParent & Kt & "," & "image" & "=" & Kt & image & Kt & "," & "link_sheet" & "=" & Kt & linkSheet & Kt & " WHERE " & "id_menu" & "=" & idMenu & ""
  
    Connect (SQLStr)
End Function

Function serverDeleteMenu(idMenu As String)
    Dim SQLStr As String
    Dim Kt As String 'Khai bao ky tu nhay don "'"
    Kt = "'"
    
    SQLStr = "DELETE FROM menu WHERE id_menu =" & idMenu & ""
    Connect (SQLStr)
End Function

