VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "The Terrorism alert status is:"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Second 
      Interval        =   1000
      Left            =   2760
      Top             =   240
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5400
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "www.whitehouse.gov"
      URL             =   "http://www.whitehouse.gov/homeland/"
      Document        =   "/homeland/"
   End
   Begin VB.Label ThreatLevel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "LOADING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Threat level is being loaded."
      Height          =   15
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SupressMessage As Boolean
Dim A As Integer ' Clock ticks, Set on a 5 minute update period.

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    
    Label1 = Inet1.OpenURL ' The URL is set at design time to 'www.whitehouse.gov/homeland'
    
    If Label1 = "" Then ' did not load the site
        If Not SupressMessage Then MsgBox "Terror alerter could not connect or there was an error."
        SupressMessage = True
        Exit Sub
    End If
    
    Label2 = Right(Label1, Len(Label1) - InStr(1, Label1, "Last updated:") + 1) 'Find the 'Last updated:' section, Put it into label2
    Label2 = Left(Label2, InStr(1, Label2, "</font>") - 1) 'Strip the tailing HTML from the 'Last updated:' section.

If InStr(1, UCase(Label1), "GUARDED") Then
    ThreatLevel = "GAURDED"
    ThreatLevel.ForeColor = vbGreen
End If

If InStr(1, UCase(Label1), "LOW") Then
    ThreatLevel = "LOW"
    ThreatLevel.ForeColor = vbBlue
End If

If InStr(1, UCase(Label1), "ELEVATED") Then
    ThreatLevel = "ELEVATED"
    ThreatLevel.ForeColor = vbYellow
End If

If InStr(1, UCase(Label1), "HIGH") Then
    ThreatLevel = "HIGH"
    ThreatLevel.ForeColor = &H80FF&
End If

If InStr(1, UCase(Label1), "SEVERE") Then
    ThreatLevel = "SEVERE"
    ThreatLevel.ForeColor = vbRed
    MsgBox "Severe terror alert!, If i were you, I would turn on the news."
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Please, Vote for this code if you liked it."
End Sub

Private Sub Second_Timer()
    A = A + 1
    If A = 300 Then
        Form_Load
        Me.Caption = "Updated at: " & Time
        A = 0
    End If
End Sub

