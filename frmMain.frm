VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB WebCam V2 ( Mohammed Samir Fayed )"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3810
      Width           =   1845
   End
   Begin VB.CommandButton cmdSavePic 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Save Picture"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3810
      Width           =   1845
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3810
      Width           =   1245
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3810
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Height          =   3660
      Left            =   3660
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   60
      Width           =   4860
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   5370
      Top             =   3840
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' /////////////////////////////////////////////////
'
'            Mohammed Samir Fayed
'
'     View and Capture WebCam Picture
'     WITHOUT using ClipBoard , Timer ,OCX ,DLL
'
' ////////////////////////////////////////////////


Option Explicit



Private Sub cmdSavePic_Click()
 If Dir(App.Path & "\myPic", vbDirectory) = "" Then MkDir (App.Path & "\myPic")
    Set Me.Image1.Picture = hDCToPicture(GetDC(modWebCam.hHwnd), 0, 0, 320, 240)
    
    SavePicture Image1.Picture, App.Path & "\myPic\" & Format(Date, "dd_MM_yyyy") & " " & Format(Time, "hhmmss") & ".bmp"
End Sub

Private Sub cmdStart_Click()
    ' Srart WebCam Capture
    
    If Me.List1.ListCount = 0 Then Exit Sub
    If List1.ListIndex = -1 Then Exit Sub
    
    modWebCam.OpenPreviewWindow List1.ListIndex, Me.Picture1
    
End Sub


Private Sub cmdStop_Click()
' Stop Capture
    modWebCam.ClosePreviewWindow
End Sub

Private Sub Command3_Click()
    modWebCam.ClosePreviewWindow
    Unload Me
End Sub

Private Sub Form_Load()
    modWebCam.LoadDeviceList Me.List1
End Sub


