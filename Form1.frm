VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "API Demo "
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Left            =   5685
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Left            =   5745
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      Height          =   2130
      Left            =   2490
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2070
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   885
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2130
      Left            =   105
      Picture         =   "Form1.frx":515D
      ScaleHeight     =   2070
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   855
      Width           =   2295
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   30
      ScaleHeight     =   1755
      ScaleWidth      =   7545
      TabIndex        =   1
      Top             =   3105
      Width           =   7605
      Begin VB.Label l6 
         BackStyle       =   0  'Transparent
         Caption         =   "API"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   795
         Width           =   1830
      End
      Begin VB.Label l7 
         BackStyle       =   0  'Transparent
         Caption         =   "AlphaBlend"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   645
         Left            =   60
         TabIndex        =   4
         Top             =   1020
         Width           =   7470
      End
      Begin VB.Label l5 
         BackStyle       =   0  'Transparent
         Caption         =   "This demo will show you how to use new AlphaBlend function to give semi transperent effect."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   60
         TabIndex        =   2
         Top             =   315
         Width           =   7470
      End
      Begin VB.Label l4 
         BackStyle       =   0  'Transparent
         Caption         =   "Example Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   45
         Width           =   1830
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10000
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "Form1.frx":F49E
         Top             =   120
         Width           =   4140
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const AC_SRC_OVER = &H0

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Function AlphaBlend Lib "msimg32.dll" _
    (ByVal hdc As Long, _
    ByVal lInt As Long, _
    ByVal lInt As Long, _
    ByVal lInt As Long, _
    ByVal lInt As Long, _
    ByVal hdc As Long, _
    ByVal lInt As Long, _
    ByVal lInt As Long, _
    ByVal lInt As Long, _
    ByVal lInt As Long, _
    ByVal BLENDFUNCT As Long) As Long
    
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" _
    (Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

Dim BlendVal As Integer
Dim BF As BLENDFUNCTION, lBF As Long

'//Variable to hold images so we can swap
Dim tempPic1  As New StdPicture, tempPic2 As New StdPicture

Private Sub Form_Load()
        
    '//Load images from file
    'Picture2.Picture = LoadPicture(App.Path & "\" & "ash01.jpg")
    'Picture1.Picture = LoadPicture(App.Path & "\" & "ash02.jpg")
    
    'Set the graphics mode to persistent
    Picture1.AutoRedraw = True
    Picture2.AutoRedraw = True
    
    'API uses pixels
    Picture1.ScaleMode = vbPixels
    Picture2.ScaleMode = vbPixels
    
    'set the parameters
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    BlendVal = 1
    Set tempPic1 = Picture1.Picture
    Set tempPic2 = Picture2.Picture
    
    Timer1.Interval = 200
    Timer2.Interval = 200
    
    Timer1.Enabled = True
    Timer2.Enabled = False
        
End Sub

Public Sub DoAlphablend(SrcPicBox As PictureBox, DestPicBox As PictureBox, AlphaVal As Integer)
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = AlphaVal
        .AlphaFormat = 0
    End With
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend DestPicBox.hdc, 0, 0, DestPicBox.ScaleWidth, DestPicBox.ScaleHeight, SrcPicBox.hdc, 0, 0, SrcPicBox.ScaleWidth, SrcPicBox.ScaleHeight, lBF
End Sub
Private Sub Timer1_Timer()
    Picture1.Refresh
    Picture2.Refresh

    BlendVal = BlendVal + 5
    If BlendVal >= 155 Then
        flag = True
        Timer1.Enabled = False
        Picture2.Picture = tempPic1
        Timer2.Enabled = True
        BlendVal = 1
    End If
    
    DoAlphablend Picture2, Picture1, BlendVal

    Me.Caption = CStr(BlendVal)
End Sub
Private Sub Timer2_Timer()
    Picture1.Refresh
    Picture2.Refresh

    BlendVal = BlendVal + 5
    If BlendVal >= 155 Then
        BlendVal = 1
        Timer1.Enabled = True
        Timer2.Enabled = False
        Picture2.Picture = tempPic2
    End If
    DoAlphablend Picture2, Picture1, BlendVal

    Me.Caption = CStr(BlendVal)
End Sub
