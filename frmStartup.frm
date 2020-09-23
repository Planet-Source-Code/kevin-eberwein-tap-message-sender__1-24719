VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmStartup 
   Caption         =   "TAP Message Sender"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDial9 
      Caption         =   "Dial 9 for an outside line."
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtModemPort 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Text            =   "1"
      Top             =   480
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   6600
      Max             =   16
      Min             =   1
      TabIndex        =   3
      Top             =   480
      Value           =   1
      Width           =   255
   End
   Begin VB.ListBox lstChecks 
      Height          =   2205
      Left            =   480
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton cmdOne 
      Caption         =   "CheckSum"
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtPin 
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtMessage 
      Height          =   1695
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   5415
   End
   Begin VB.TextBox txtPhoneNumber 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6960
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "Modem Comm Port:"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Pin Number:"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Message:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'*
'* TAP Message Sender
'*
'* Written: July 3, 2001
'* By: Kevin Eberwein (keverwein@nc.rr.com)
'*
'* Purpose: This application will send alphanumeric pages using a modem.
'*
'* Requirements: A modem, an alphanumeric pager, the phone number you paging
'*  service uses for sending alpha pages, and your pin number.
'*
'****************************************************************************

Const nMaxConnectWait = 60 'Wait up to 60 Seconds for the connection
Const t1 = 2, t2 = 1, t3 = 10, t4 = 4, t5 = 8 'Timing parameters as defined by TAP v1.8

Private Sub cmdOne_Click()
  MsgBox TAPCheckSum(Chr(2) + Trim(txtPin.Text) + vbCr + Trim(txtMessage.Text) + vbCr + Chr(3))
End Sub

Private Sub cmdSend_Click()
  Dim nResults As Integer, sMsg As String, sTemp As String
  
  Screen.MousePointer = vbHourglass
  
  sTemp = txtMessage.Text
  'Split the message into 200 character chunks to send
  While Len(sTemp) > 200
    sMsg = Left(sTemp, 200)
    sTemp = Right(sTemp, Len(sTemp) - 200)
    nResults = SendTAPPage(sMsg, txtPin.Text, txtPhoneNumber.Text)
    Select Case (nResults)
      Case -1: sMsg = "Failure in Tap startup code."
      Case 0: sMsg = "Successful Completion."
      Case 1: sMsg = "Step 1 Failure."
      Case 2: sMsg = "Step 2 Failure."
      Case 3: sMsg = "Step 3 Failure."
      Case 4: sMsg = "Step 4 Failure."
      Case 5: sMsg = "Step 5 Failure."
      Case 6: sMsg = "Step 6 Failure."
      Case 7: sMsg = "Step 7 Failure."
      Case 8: sMsg = "Step 8 Failure."
      Case -8: sMsg = "Step -8 Failure."
      Case 9: sMsg = "Step 9 Failure."
      Case 10: sMsg = "Step 10 Failure."
    End Select
    If nResults <> 0 Then
      MsgBox sMsg
    End If
  
  Wend
  nResults = SendTAPPage(sTemp, txtPin.Text, txtPhoneNumber.Text)
  
  Select Case (nResults)
    Case -1: sMsg = "Failure in Tap startup code."
    Case 0: sMsg = "Successful Completion."
    Case 1: sMsg = "Step 1 Failure."
    Case 2: sMsg = "Step 2 Failure."
    Case 3: sMsg = "Step 3 Failure."
    Case 4: sMsg = "Step 4 Failure."
    Case 5: sMsg = "Step 5 Failure."
    Case 6: sMsg = "Step 6 Failure."
    Case 7: sMsg = "Step 7 Failure."
    Case 8: sMsg = "Step 8 Failure."
    Case -8: sMsg = "Step -8 Failure."
    Case 9: sMsg = "Step 9 Failure."
    Case 10: sMsg = "Step 10 Failure."
  End Select
  
  MsgBox sMsg
  
  Screen.MousePointer = vbDefault
  
End Sub

Function TAPCheckSum(sInString As String) As String
  'Compute the TAP checksum value for any string sent in
  Dim sResults As String, nCheckSum As Long
  Dim I As Integer
  Dim n8 As Integer, n4 As Integer, n1 As Integer
  Dim n As Integer
  
  sResults = ""
  nCheckSum = 0
  lstChecks.Clear
  'First, add all the ascii values of each character in the string.
  For I = 1 To Len(sInString)
    lstChecks.AddItem Mid(sInString, I, 1) + " = " + Trim(Str(Asc(Mid(sInString, I, 1))))
    nCheckSum = nCheckSum + Asc(Mid(sInString, I, 1))
  Next I
  
  lstChecks.AddItem "nCheckSum = " + Trim(Str(nCheckSum))
  
  'If the check sum is greater than 4096, 1000000000000 in Binary, remove all bits so
  ' only the right 12 are available
  If nCheckSum >= 4096 Then
    nCheckSum = nCheckSum - Int(nCheckSum / 4096) * 4096
  End If
  
  'Now, take the result and get the 3 sets of 4 bits
  n8 = Shri(nCheckSum, 8)
  n4 = Shri((nCheckSum - Shli(n8, 8)), 4)
  n1 = nCheckSum - Shli(n8, 8) - Shli(n4, 4)
  
  lstChecks.AddItem "n8 = " + Trim(Str(n8)) + ", n4 = " + Trim(Str(n4)) + ", n1 = " + Trim(Str(n1))

  'Add 48 to each number to get the numbers to be the ascii code of the checksum bitset.
  n1 = n1 + 48
  n4 = n4 + 48
  n8 = n8 + 48
  lstChecks.AddItem "n8 = " + Trim(Str(n8)) + ", n4 = " + Trim(Str(n4)) + ", n1 = " + Trim(Str(n1))
  lstChecks.AddItem ":" + Chr(n8) + Chr(n4) + Chr(n1) + ":"
  
  'Set the checksum equal to the characters represented by the bits
  sResults = Chr(n8) + Chr(n4) + Chr(n1)
  
  TAPCheckSum = sResults
End Function

Function SendTAPPage(sInMsg As String, sPin As String, sPhone As String) As Integer
  'This is where all the code for sending the pages over the modem resides.
  ' If the program fails in any step, we stop processing and do a hard disconnect.
  Dim DialStr As String, bStepDone As Boolean
  Dim StartTime As Date, bOK As Boolean
  Dim nCnt As Long, sFromModem As String
  Dim sMsg1 As String, sMsg2 As String
  Dim nResult As Integer, bTryAgain As Boolean
  
  'Set nResult to the step number we are on.  If the page goes, set it to 0.
  nResult = -1
  
  ' Communications port settings.
  MSComm1.CommPort = Val(txtModemPort.Text)
  MSComm1.Settings = "9600,E,7,1"
  
  'Setup the dialing string
  DialStr = "ATDT"
  If chkDial9.Value = vbChecked Then
    DialStr = DialStr + "9,"
  End If
  DialStr = DialStr + Trim(sPhone) + vbCr
  
  'Generate the Message string to send in TAP Step 8.  This is done here to
  ' save time durring the sending of the page.
  sMsg1 = Chr(2) + Trim(sPin) + vbCr + Trim(sInMsg) + vbCr + Chr(3)
  sMsg2 = sMsg1 + TAPCheckSum(sMsg1) + vbCr
  
  ' Open the communications port.
  On Error Resume Next
  MSComm1.PortOpen = True
  If Err Then
     MsgBox "COM" + Trim(txtModemPort.Text) + ": not available. Change the CommPort property to another port."
     Exit Function
  End If
  
  'Since the I am using a network modem, I need to wait 5 seconds to make
  ' sure it is connected.  Comment this out if your modem is on a local com port.
  StartTime = Now
  While Abs(DateDiff("s", StartTime, Now)) < 5
    DoEvents
  Wend
  
  ' Flush the input buffer.
  MSComm1.InBufferCount = 0
  
  ' TAP Step 1: Dial the number.
  nResult = 1
  MSComm1.Output = DialStr
  StartTime = Now
  
  ' TAP Step 2: Wait for the connect message
  '   (TAP says wait for Carrier Up signal, it's the same thing in the end.)
  nResult = 2
  bStepDone = False
  bOK = False
  nCnt = 0
  While Not bStepDone
    DoEvents
    ' If there is data in the buffer, then read it.
    If MSComm1.InBufferCount > 0 Then
      sFromModem = LCase(MSComm1.Input)
      ' Check for "Connect".
      If InStr(sFromModem, "connect") <> 0 Then
        ' We are connected
        bStepDone = True
        bOK = True
      End If
    Else
      If Abs(DateDiff("s", StartTime, Now)) >= nMaxConnectWait Then 'Wait for the connection based on the nMaxConnectWait constant
        bStepDone = True
        bOK = False
      End If
    End If
  Wend
  
  If bOK Then 'We are connected.  Send a vbCr once a second until t5 seconds have
              'gone by or we receive an "ID=" from the Pagining Terminal
    ' TAP Step 3: Send <CR> until Step 4 occurs
    nResult = 3
    MSComm1.Output = vbCr
    StartTime = Now
    bOK = False
    bStepDone = False
    nCnt = 0
    nResult = 4
    While Not bStepDone
      DoEvents
      If MSComm1.InBufferCount > 0 Then
        sFromModem = LCase(MSComm1.Input)
        ' TAP Step 4: Check for "ID=".
        If InStr(sFromModem, "id=") <> 0 Then
          ' We are connected
          bStepDone = True
          bOK = True
        End If
      Else
        If Abs(DateDiff("s", StartTime, Now)) >= 1 Then '1 or more seconds have gome by
          StartTime = Now
          MSComm1.Output = vbCr
        End If
        If nCnt >= t5 Then  'Wait t5 seconds or so for Step 4 to occur.
          bStepDone = True
          bOK = False
        End If
      End If
    Wend
  End If
  
  If bOK Then 'We have the next step of TAP.
    ' TAP Step 5: Send "<ESC>SST"  See TAP for SST definition.
    nResult = 5
    MSComm1.Output = Chr(27) + "PG1" + vbCr
    
    bTryAgain = True
    While bTryAgain
      bTryAgain = False
      ' TAP Step 6: Look for a response to step 5
      '  - Note: we should see and ACK, NAK, or EOT (End of Transmission)
      nResult = 6
      StartTime = Now
      bOK = False
      bStepDone = False
      nCnt = 0
      While Not bStepDone
        DoEvents
        If MSComm1.InBufferCount > 0 Then
          sFromModem = LCase(MSComm1.Input)
          If InStr(sFromModem, Chr(6)) <> 0 Then
            ' We have received the ACK
            bStepDone = True
            bOK = True
          End If
          If InStr(sFromModem, Chr(21)) <> 0 Then
            ' We have received the NAK.  We should try again
            bStepDone = True
            bTryAgain = True
          End If
          If InStr(sFromModem, Chr(27) + Chr(4)) <> 0 Then
            ' We have received the EOT. Pager Teminal is telling us to quit.
            bStepDone = True
            bOK = False
          End If
        Else
          If Abs(DateDiff("s", StartTime, Now)) >= 1 Then 'Wait for ACK, NAK, or EOT
            nCnt = nCnt + 1
            StartTime = Now
          End If
          If nCnt >= t3 Then 't3 seconds to wait for reply
            bStepDone = True
            bOK = False
          End If
        End If
      Wend
    Wend
  End If
  
  If bOK Then
    ' TAP Step 7: Wait for the go ahead from the Pager Terminal
    nResult = 7
    
    'See if we got the go ahead in the last message.
    If InStr(sFromModem, Chr(27) + "[p") = 0 Then
      StartTime = Now
      bOK = False
      bStepDone = False
      nCnt = 0
      While Not bStepDone
        DoEvents
        If MSComm1.InBufferCount > 0 Then
          sFromModem = LCase(MSComm1.Input)
          If InStr(sFromModem, Chr(27) + "[p") <> 0 Then
            ' We have received the go ahead
            bStepDone = True
            bOK = True
          End If
        Else
          If Abs(DateDiff("s", StartTime, Now)) >= 1 Then 'Wait for go ahead
            nCnt = nCnt + 1
            StartTime = Now
          End If
          If nCnt >= t3 Then 't3 seconds to wait for reply
            bStepDone = True
            bOK = False
          End If
        End If
      Wend
    End If
  End If
  
  If bOK Then
  
    bTryAgain = True
    While bTryAgain
      bTryAgain = False
      ' TAP Step 8: Now, send the pin, the text, and then the checksum
      '  (This was determined at the start of the function to save time.
      '   The protocol allows for this to be sent in t3 seconds after the last step.)
      nResult = 8
      MSComm1.Output = sMsg2
      StartTime = Now
      bOK = False
      bStepDone = False
      nCnt = 0
      nResult = -8 '(Let's us know that we are waiting in step 8b since it's a two part step)
      While Not bStepDone
        DoEvents
        If MSComm1.InBufferCount > 0 Then
          sFromModem = LCase(MSComm1.Input)
          ' Look for ACK, NAK, RS, or EOT
          If InStr(sFromModem, Chr(6)) <> 0 Then
            ' We have received the ACK
            bStepDone = True
            bOK = True
          End If
          If InStr(sFromModem, Chr(30)) <> 0 Then
            ' We have received the RS.  We should try again.
            bStepDone = True
            bTryAgain = True
          End If
          If InStr(sFromModem, Chr(21)) <> 0 Then
            ' We have received the NAK.  We should try again.
            bStepDone = True
            bTryAgain = True
          End If
          If InStr(sFromModem, Chr(27) + Chr(4)) <> 0 Then
            ' We have received the EOT. Pager Teminal is telling us to quit.
            bStepDone = True
            bOK = False
          End If
        Else
          If Abs(DateDiff("s", StartTime, Now)) >= 1 Then 'Wait for ACK, NAK, or force disconnect
            nCnt = nCnt + 1
            StartTime = Now
          End If
          If nCnt >= t3 Then 't3 seconds to wait for reply
            bStepDone = True
            bOK = False
          End If
        End If
      Wend
    Wend
  End If
  
  If bOK Then
    ' TAP Step 9: Send the End of Tranmission Code (EOT)
    nResult = 9
    MSComm1.Output = Chr(4) + vbCr
    StartTime = Now
    bOK = False
    bStepDone = False
    nCnt = 0
    nResult = 10
    While Not bStepDone
      DoEvents
      If MSComm1.InBufferCount > 0 Then
        ' TAP Step 10: Receive any messages or error messages prior to the disconnect.
        sFromModem = sFromModem + LCase(MSComm1.Input)
        If InStr(sFromModem, Chr(4)) <> 0 Or InStr(sFromModem, "no car") <> 0 Then
          'The paging terminal has ended the transmission
          bStepDone = True
          bOK = True
        End If
      Else
        If Abs(DateDiff("s", StartTime, Now)) >= 60 Then 'Wait 60 seconds for all data to come in
          bStepDone = True
          bOK = True
        End If
      End If
    Wend
  End If
  
  
  ' Disconnect the modem.
  MSComm1.Output = "+++ATH" + vbCr
  
  ' Close the port.
  MSComm1.PortOpen = False

  If bOK Then
    nResult = 0
  End If
  
  TAPPage = nResult

End Function

Private Sub VScroll1_Change()
  txtModemPort.Text = Trim(Str(VScroll1.Value))
End Sub
