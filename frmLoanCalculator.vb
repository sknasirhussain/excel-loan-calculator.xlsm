Option Explicit

Private Sub cmdCalculate_Click()
    
    Application.DisplayAlerts = False
    
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set conn = New ADODB.Connection
    conn.Open "driver={sql server};server=LINKOLFEB24-092;database=db;uid=admin;pwd=admin123;"
    Debug.Print "Connected successfully"
    
    Dim years As Long
    Dim duration As Long, rate As Double
    Dim sql As String
    
    duration = 0
    
    years = tbxDuration.Value
    
    If years <= 1 Then
        duration = 1
    ElseIf years >= 2 And years <= 5 Then
        duration = 5
    ElseIf years >= 6 Then
        duration = 10
    End If
    
    sql = "Select Rate from rates where Duration =" & duration
    Set rs = conn.Execute(sql)
    Debug.Print rs.Fields("Rate").Value
    rate = rs.Fields("Rate").Value / 100
    
    Cells(4, 4).Value = CLng(tbxAmount)
    Cells(5, 4).Value = rate
    Cells(6, 4).Value = CLng(tbxDuration)
    Cells(8, 4).Value = CLng(tbxNumberOfPayments)
    
    Application.DisplayAlerts = True
    
    Unload Me
End Sub
