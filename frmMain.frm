VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCPConn"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wsk_conn 
      Left            =   240
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton btn_send 
      Caption         =   "&Send"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox txt_send 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   5655
   End
   Begin VB.TextBox txt_incoming 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   4455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   1440
      Width           =   6615
   End
   Begin VB.Frame fra_connectiondetails 
      Caption         =   "Connection Details:"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "?"
         Height          =   255
         Left            =   6000
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton btn_disconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton btn_connect 
         Caption         =   "&Connect"
         Default         =   -1  'True
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_port 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txt_address 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbl_status 
         BackColor       =   &H00000000&
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   4920
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lbl_status_status 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lbl_port 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Port:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl_address 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Address:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label lbl_send 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Incoming Data from Server:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   6615
   End
   Begin VB.Label lbl_incoming 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Incoming Data from Server:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   6615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_connect_Click()
' Actually connect the the address:port
TCPConnect frmMain.txt_address, frmMain.txt_port
End Sub
Private Sub btn_disconnect_Click()
With frmMain.wsk_conn
'Close Winsock and disconnect from the address
.Close
End With
MsgBox "Connection closed.", vbInformation, "Message"
' Set everything up again
SetupTCPClosureEnvironment
End Sub
Private Sub btn_send_Click()
' Pass text to the send function
TCPSend frmMain.txt_send.Text
' Blank the input box
frmMain.txt_send = ""
End Sub
Private Sub Command1_Click()
' Load the form
Load frmAbout
' Show the form on screen (but modally!)
frmAbout.Show vbModal
End Sub
Private Sub Command2_Click()
' Say our farewells
MsgBox "Thank you for using my project - Hope it was of some use! Good bye - Craig Malton.", vbInformation, "Message"
' Close winsock if it's open
frmMain.wsk_conn.Close
' Unload forms
Unload frmAbout
Unload frmMain
' . . . and quit
End
End Sub
Private Sub wsk_conn_Close()
' Inform user of closure
MsgBox "Connection closed.", vbInformation, "Message"
' Set everything up again
SetupTCPClosureEnvironment
End Sub
Private Sub wsk_conn_Connect()
' Inform user of connection
MsgBox "Connection established.", vbInformation, "Message"
' Set everything to be usable
SetupTCPConnectEnvironment
End Sub
Private Sub wsk_conn_DataArrival(ByVal bytesTotal As Long)
' Actually GET the data from the socket
frmMain.wsk_conn.GetData incomingdata, vbString
' Display it on screen
frmMain.txt_incoming.SelText = vbCrLf & incomingdata
End Sub
Private Sub wsk_conn_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' Inform user that there has been a Winsock problem
MsgBox "There has been an error involving Winsock - This error is:" & vbCrLf & Description & vbCrLf & "Error code:" & vbCrLf & Number, vbInformation, "Message"
End Sub
