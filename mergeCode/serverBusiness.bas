Attribute VB_Name = "serverBusiness"
'Lay danh sach ke hoach danh thu
Function getBusinessPlan()
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim rs As ADODB.RecordSet
    Set rs = New ADODB.RecordSet
    Dim k As Long
    Dim Row As Integer
    Dim rowEnd As Long 'Dong cuoi co du lieu
    Dim sheetDataUnit As Worksheet 'Sheet DataKe hoach doanh thu dong vi
    Set sheetDataUnit = Sheet27
    StrCnn = "Driver={SQL Server};Server=192.168.100.72,1433;Database=bos_vba;Uid=bosdev;Pwd=bos123456;"
 
    'Xu ly lenh
    Dim SQLStr As String
    SQLStr = "SELECT * FROM Ds_KheHoachKinhDoanhDonVi"
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    rs.Open SQLStr, Cn, adOpenStatic
        'Ten thTable
        For Each Field In rs.Fields
            sheetDataUnit.Range("A2").Offset(0, k).value = Field.name
            sheetDataUnit.Range("A2").Offset(0, k).Borders.LineStyle = True
            sheetDataUnit.Range("A2").Offset(0, k).Font.Color = vbWhite
            sheetDataUnit.Range("A2").Offset(0, k).Interior.Color = vbBlue
            k = k + 1
        Next Field
        sheetDataUnit.Range("A3").CopyFromRecordset rs
        
        rowEnd = sheetDataUnit.Range("A" & Rows.Count).End(xlUp).Row
        
        For Row = 2 To rowEnd
            sheetDataUnit.Range("A" & Row).RowHeight = 30
        Next Row
        
    
    Cn.Close
    Set Cn = Nothing
    
    
End Function

Function serverCreateBusinessPlan(valueUnitID As String, valueYear As String, valueRevenuePlan As String, valueActualRevenue As String)
    Dim SQLStr As String
    Dim resultTD As String
    Dim Kt As String 'Khai bao ky tu nhay don "'"
    Kt = "'"
        
    If valueActualRevenue = "" Then
        resultTD = 0
    Else
        resultTD = valueActualRevenue
    End If
    
    
    'Xu ly lenh

    SQLStr = "INSERT INTO KeHoachDoanhThu(PhongBanID, Nam, KeHoachDoanhThu) VALUES(" & valueUnitID & "," & valueYear & "," & valueRevenuePlan & ") ; INSERT INTO DoanhThuThucDat(PhongBanID, Nam, DoanhThuThucDat) VALUES(" & valueUnitID & "," & valueYear & "," & resultTD & ")"
    Debug.Print SQLStr
    Connect (SQLStr)
End Function

Function serverGetBusinessPlanYear(valueYear As String) As Variant
    Dim result As Variant
    Dim SQLStr As String
    
    SQLStr = "SELECT KeHoachDoanhThuID,k.PhongBanID,k.Nam,KeHoachDoanhThu,DoanhThuThucDat FROM KeHoachDoanhThu k LEFT JOIN DoanhThuThucDat t ON k.Nam = t.Nam Where k.Nam =" & valueYear & ""
    result = sqlGetRows(SQLStr)
    
    serverGetBusinessPlanYear = result
End Function

