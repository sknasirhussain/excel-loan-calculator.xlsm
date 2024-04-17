Option Explicit

Sub FetchRate()

    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set conn = New ADODB.Connection
    conn.Open "driver={sql server};server=LINKOLFEB24-092;database=db;uid=admin;pwd=admin123;"
    Debug.Print "Connected successfully"
    
    Dim rng As Long
    Dim duration As Long, rate As Double
    Dim sql As String
    
    duration = 0
    
    rng = Cells(6, 4).Value
    
    If rng <= 1 Then
        duration = 1
    ElseIf rng > 2 And rng <= 5 Then
        duration = 5
    ElseIf rng > 6 Then
        duration = 10
    End If
    
    sql = "Select Rate from rates where Duration =" & duration
    Set rs = conn.Execute(sql)
    Debug.Print rs.Fields("Rate").Value
    rate = rs.Fields("Rate").Value / 100
    Cells(5, 4).Value = rate
 
End Sub

