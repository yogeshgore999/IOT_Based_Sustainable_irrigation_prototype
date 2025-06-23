Dim cktb As Boolean
Dim gsmb As Boolean
Dim X As Boolean 'to track serial port on/
Dim temp As Boolean
Dim t1_data As Byte
Dim hum As Boolean
Dim h1_data As Byte
Dun n1 As Integer
Dim illu As Boolean
Dim illu1 As Boolean
Dim il_data As Byte
Dim iil_data As Byte
Dim lpg As Boolean
Dim 11_data As Byte
Dim oXL As Excel.Application
Dim oWB As Excel Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel. Range
Private blnConnected As Boolean
Private Sub cmdSend_Click()
Call post data(txtUrl.Text, "status", "A")
End Sub
Function post data(ByVal i_url As String, ByVal v name As String, ByVal v value As String)
Dim eUrl As URL
Dim strMethod As String
Dim strData As String
Dim strPostData As String
Dim strHeaders As String
Dim striITT As String
Dim X As Integer
strPostData= "”
strHeaders= “”
strMethod cboRequestMethod. List(cboRequestMethod ListIndex)
If blnConnected Then Exit Function
get the url
eUrl=ExtractUrl(i_url)
If eUrl.Host = vbNullString Then
MsgBox "Invalid Host", vbCritical, "ERROR"
Exit Function
End If
strDataLeft(strData, Len(strData) - 1)
If strMethod="GET" Then
if this is a GET request then the URL. encoded data
is appended to the URI with a ?
If eUrl. Query <> vbNullString Then
eUrL.URI eUrl URI & "&" & strData
Else
eUrL.URI = eUrl URI && strData
End If
Else
if it is a post request, the data is appended to the
body of the HTTP request and the headers Content-Type
and Content-Length added
strPostData = strData
strHeaders = "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
"Content-Length: " & Len(strPostData) & vbCrLf
End If
End If
get any additional headers and add them
For X-0 To txtHeaderName.Count - 1
If txtHeaderName(X). Text <> vbNullString Then
strHeaders = strHeaders & txtHeader Name(X).Text & " " &
txtHeaderValue(X).Text & vbCrLf
End If
Next X
clear the old HTTP response
txtResponse.Text= “”
build the HTTP request in the form
{REQ METHOD} URI HTTP/1.0
Host: {bost}
{headers}
{post data}
strHTTP = strMethod & “ “& eURL & “HTTP/1.0” & vbCrLf
strHTTP = strHTTP & "Host:& eUrl.Host & vbCrLf
strHTTP = strHTTP & strHeaders
strHTTP = strHTTP & vbCrLf
strHTTP = HTTP & simPostData
xiRequest.Text = strHTTP
Winsock.Connect
wait for a connection
While Not blnConnected
DoEvents
Wend
send the HTTP request
Winsock.SendData strHTTP
wack.SendDats "http://Domain name[.] com” seni u php status-S

Private Sub cmdcircuit_Click()
On Error GoTo com error
If cktb False Then
Cktb = True
cmdcircuit.Caption = "Disconnect”
MSComm1.CommPort = Combo1.Text
MSComm1.PortOpen = True
Else
Ckib = False
cmdcircuit.Caption = "Connect"
MSComm1.PortOpen = False
End If
Exit Sub
com)error:
MsgBox "Another program is using serial port, try again after other program completes", vbCritical,
"Communication Error"
Cktb = False
cmdcircuit.Caption = "Connect"
End Sub
Private Sub Command2_Click()
MSComm1.Output = "W"
MSComm1.Output = "W"
MSComm1.Output = "W"
MSComm1.Output = "X"
End Sub
Private Sub Timer4_Timer()
Call post_data(txtUrl. Text, "status", Iblcv.Caption)
Call delay(2)
Call post_data(txtUrl1.Text, "status", Iblhum.Caption)
Cell vdelay(2)
Call post_data(txtUrl2.Text; "status", “”)
End Sub
Private Sub Timer5_Timer()
If cktb = True Then
If InStr(txtResponse Text. "QQ") 0 Then
Label 12.Caption = "ON"
If cktb = True Then
MSComm1.Output "T"
Call vdelay(1)
MSComm1.Output = "X"
End If
Elself InStritxtResponse Text, "WW")> 0 Then
Label12.Caption"OFF"
If cktb = True Then
MSComm1. Output = “Y”
Call vdelay(1)
MSComm1.Ouput= “X”
End If
End If
End If
End Sub
Private Sub winock_Connect()
blaConnected = True
End Sub
This event occurs when data is arriving via winsock
Private sub winsock_DataArrival(By Val bytesTotal as Long)
Dim strResponse As String
Winsock.GetData Str Response, vbString,bytesTotal
strResponse = FormatLineEndings(strResponse)
We append this to the response box because data arrives in Multipe Packets
txtResponse.Text = txtResponse.Text & strResponse
End Sub

Private Sub winsock Errort ByVal Number As Integer. Description As String, ByVal Scode As Long. ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, Cancel Display As Boolean)
MagBox Description, vbExclamation, "ERROR"
Winsock.Close
End Sub
Private Sub winsock_Close()
binConnected = False
Winsock.Close
End Sub
this function converts all line endings to Windows CrLf line endings
Private Function FormatLineEndings(ByVal str. As String) As String
Dim prevChar As String
Dim nextChar As String
Dim curChar As String
Dim strRet As String
Dim X As Long
prevChar = “”
nextChar = “”
curChar = “”
strRet = “”
For X = 1 To Len(str)
prevChar = curChar
curChar Mid$(str, X, 1
If nextChar <> vbNullString And curChar <> nexiChar Then
curChar = curChar & nextChar
nextChar = ""
Elself curChar = vbL.f Then
If prevChar <> vbCr Then
curChar = vbCrL.f
End If
nextChar = "”
Elself curChar = vbCr Then
nextCharvbl.f
End If
strRet = strRet & curChar
Next X
FormatLineEndings = strRet
illudata False
Ipg-False
Ipgdata False
MSComm1 Settings"9600N,8,1
MSComm1 PortOpen =True
Cktb =False
Gsmb = False
Combo1.AddItem “1”
Combo1.AddItem “2”
Combo1.AddItem “3”
Combo1.AddItem “4”
Combo1.AddItem “5”
Combo1.AddItem “6”
Combo1.AddItem “7”
Combo1.AddItem “8”
Combo1.AddItem “9”
Combo1.AddItem “10”
End Sub
Private Sub vdelay(ByVal How Long As Date)
Dim endDate As Date
endDate Date Add("s", How Long, Now)
While endDate > Now
DoEvents 'Allows windows to handle other stuff
Wend
End Sub
Public Function Random Number(ByVal Max Value As Long, Optional
ByVal MinValue As Long=0)
On Error Resume Next
Randomize Timer
Iblheart.Caption Int(MaxValue MinValue 1)* Rnd) + MinValue
End Function
Private Sub Form_Unload(Cancel As Integer)
MSComm1.PortOpen = False
oWB.Save
oXL Quit
Set oRng = Nothing
Set oSheet = Nothing
Set oWB = Nothing
Set OXL = Notlung
end
End Sub
Private Sub Timer1_Timer()
If MSComm1.InBufferCount Then
Label1.Caption = MSComm1.Input
If Label1.Caption = "t" Then
temp = True
Elself temp = True Then
t_data Label1.Caption
If Label1.Caption = "" Then
t1. data = 0
Else
t1_data Asc(t_data)
End If
Iblev Caption = t1_data
temp False
End If
If Label1.Caption "h" Then
hum-True
Elself hum = True Then
H_data = Label1.Caption
If Label1.Caption = "" Then
h1_data = 0
Else
h1_data = Asc(h_data)
End If
H1_data = h1_data
H2_data = h1_data/2.55
Iblhum.Caption = 100 - h2 data
hum = False
End If
End If
End Sub
Private Sub Timer2_Timer()
Set oSheet oWB.Worksheets("Device 1")
oSheet.Cells(n1, 1).Value = nl - 2
oSheet.Cells(3, 10).Value = n1
oSheet.Cells(n1, 2).Value = DateS
oSheet.Cells(n1, 3).Value = Time$
oSheet.Cells(n1, 4).Value = lblcv.Caption
oSheet.Cells(n1, 5).Value = lblhum.Caption
oSheet.Cells(n1, 6).Value = lbllpg.Caption
oSheet.Cells(n1, 7).Value = Iblillu.Caption
oSheet.Cells(n1, 8).Value = Iblillu1.Caption
oSheet.Visible = True
oXL.UserControl = True
n1=n1+1
End Sub
Private Sub Timer3_Timer()
If cktb= True Then
If (Val(Iblcv.Caption) > Val(txtsp.Text)) And (Val(lblev.Caption) < 100) Then
MSComm1.Output = "E"
MSComm1.Output = "E"
Iblp1.Caption = "Fan On"
Call vdelay(0.5)
Else
MSComm1.Output = "R"
MSComm1.Output = "R"
Iblp1.Caption = "Fan Off"
Call vdelay(0.5)
End If
If (Val(Ibihum.Caption) < Val(txtspl.Text)) Then
MSComm1 Output = "0"
MSComm1.Output = "0"
Ibip2 Caption = "Pump2 ON"
Iblp2.Caption = "Pump2 ON"
Call vdelay(1)
Else
MSComm1.Output = "1"
MSComm1.Output = "1"
lblp2.Caption = "Pump2 OFF"
Call vdelay(1)
End If
End If
End Sub
