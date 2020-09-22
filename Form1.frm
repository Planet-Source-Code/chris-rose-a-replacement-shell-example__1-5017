VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   2145
   ClientWidth     =   7095
   Height          =   8670
   Left            =   60
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Top             =   1740
   Width           =   7215
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Replacement Shell Example"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   6855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This is what the line should look like when you want this shell in action : Shell= "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Read:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLOSE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   7440
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   240
      Picture         =   "Form1.frx":02AA
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   6615
   End
   Begin VB.Image Image3 
      Height          =   870
      Left            =   240
      Picture         =   "Form1.frx":9F2C
      Stretch         =   -1  'True
      Top             =   7200
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   0
      Top             =   6960
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
Shell "c:\windows\system\systray.exe"
Label5.Caption = App.Path & "\" & App.EXEName & ".exe"
End Sub

Private Sub Image1_Click()
retval = Shell("C:\windows\notepad.exe c:\windows\system.ini", 1)

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Top = 7200
Image3.Left = 240
Image1.Visible = False
Image3.Visible = True
End Sub


Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = True
Image3.Visible = False
End Sub


Private Sub Image3_Click()
retval = Shell("C:\windows\notepad.exe c:\windows\system.ini", 1)

End Sub

Private Sub Label1_Click()
retval = Shell("C:\windows\notepad.exe c:\windows\system.ini", 1)
End Sub


Private Sub Label2_Click()
retval = Shell("C:\windows\notepad.exe c:\windows\system.ini", 1)

End Sub


