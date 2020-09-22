VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   600
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3720
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   4200
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim i As Integer, h As Integer

Private Sub Form_Load()

    MSComm1.Settings = "9600,n,8,1" '"6,2,4,0"
    MSComm1.Handshaking = 0
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
    Dim sTmp As String
    If (MSComm1.InBufferCount > 0) Then
        sTmp = MSComm1.Input
    End If
OpenDB
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Con = Nothing
End Sub

Private Sub Timer1_Timer()
    Dim sTmp As String

    If (MSComm1.InBufferCount > 0) Then
        sTmp = Text1.Text
        sTmp = sTmp + MSComm1.Input
        Text1.Text = Trim(sTmp)
    End If

End Sub

' The data recieved from PABX shown in the textbox
' But what i want to do now just move only data into mdb file without minus symbol or header text
' How can i do please help me.
' Note: Alway come with header text for every start application and recieve data or new day started.
' sorry for my poor english
Public Sub OpenDB()
On Error GoTo cn_err
    Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data.mdb;Persist Security Info=False"
    Con.Open
    Exit Sub
cn_err:
    MsgBox Err.Description, vbOKOnly, App.Comments
End Sub


Public Function GetLineTotal(tb As TextBox)
Dim numLines As Integer

Dim i As Integer

numLines = 1
For i = 1 To Len(tb.Text)
If Mid(tb.Text, i, 2) = vbCrLf Then
  numLines = numLines + 1
End If
Next i

GetLineTotal = numLines
End Function

Private Sub Command2_Click()
Dim eText As String, eText1 As String, eText2 As String
Dim eText3 As String, eText4 As String, eText5 As String
Dim D As Integer, intSpace As Integer
intSpace = 0
D = 170
For h = 1 To GetLineTotal(Text1) - 3
    For i = IIf(h = 1, D, D + 1) To Len(Text1.Text)
    If Not Mid(Text1.Text, i, 1) = " " Then
        If Mid(Text1.Text, i, 2) = vbCrLf Then
            D = i
            intSpace = 0
            Exit For
        Else
            If intSpace = 0 Then
                eText = eText & Mid(Text1.Text, i, 1)
            ElseIf intSpace = 1 Then
                eText1 = eText1 & Mid(Text1.Text, i, 1)
            ElseIf intSpace >= 2 And intSpace <= 4 Then
                eText2 = eText2 & Mid(Text1.Text, i, 1)
            ElseIf intSpace = 5 Then
                eText3 = eText3 & Mid(Text1.Text, i, 1)
            ElseIf intSpace = 6 Then
                eText4 = eText4 & Mid(Text1.Text, i, 1)
            ElseIf intSpace > 6 Then
                eText5 = eText5 & Mid(Text1.Text, i, 1)
            End If
        End If
    Else
        intSpace = intSpace + 1
    End If
    Next i
    rs.Open "Select * from Table1", Con, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!Date = eText
    rs!Time = eText1
    rs!Ext = eText2
    rs!CO = eText3
    rs!DialNo = eText4
    rs!RingDur = eText5
    rs.Update
    rs.Close
    
    'Print eText & " - " & eText1 & " - " & eText2 & " - " & eText3 & " - " & eText4 & " - " & eText5
    eText = ""
    eText1 = ""
    eText2 = ""
    eText3 = ""
    eText4 = ""
    eText5 = ""
Next
Set rs = Nothing
End Sub

