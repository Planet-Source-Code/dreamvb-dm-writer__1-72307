VERSION 5.00
Begin VB.Form frmChar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Symbol"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   360
      Left            =   4770
      TabIndex        =   6
      Top             =   4140
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   6060
      TabIndex        =   5
      Top             =   4140
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      Height          =   600
      Left            =   6540
      ScaleHeight     =   540
      ScaleWidth      =   600
      TabIndex        =   3
      Top             =   30
      Width           =   660
      Begin VB.Label lblView 
         Alignment       =   2  'Center
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   525
         Left            =   45
         TabIndex        =   4
         Top             =   30
         Width           =   495
      End
   End
   Begin VB.ListBox LstChar 
      Columns         =   8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      IntegralHeight  =   0   'False
      Left            =   90
      TabIndex        =   2
      Top             =   675
      Width           =   7125
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   570
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   3165
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      Caption         =   "#0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   4260
      Width           =   240
   End
   Begin VB.Label lblFont 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Font:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   255
      Width           =   360
   End
End
Attribute VB_Name = "frmChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFont_Click()
    LstChar.Font = cboFont.Text
End Sub

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Call Unload(frmChar)
End Sub

Private Sub cmdInsert_Click()
    CharInsert = LstChar.Text
    SelFont = cboFont.Text
    ButtonPress = vbOK
    Call Unload(frmChar)
End Sub

Private Sub Form_Load()
Dim Cnt As Integer
    'Remove item
    Set frmChar.Icon = Nothing
    
    'Add Fontnames
    For Cnt = 0 To Screen.FontCount - 1
        cboFont.AddItem Screen.Fonts(Cnt)
    Next Cnt
    'Add Chars
    For Cnt = 32 To 255
        LstChar.AddItem Chr(Cnt)
    Next Cnt
    
    'Set indexs
    cboFont.ListIndex = FindInList(cboFont, "Arial")
    LstChar.ListIndex = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmChar = Nothing
End Sub

Private Sub LstChar_Click()
    lblView.Caption = LstChar.Text
    lblView.Font = cboFont.Text
    If (LstChar.ListIndex > 126) Then
        lblChar.Caption = "Character Code: Alt+0" & 32 + LstChar.ListIndex
    Else
        lblChar.Caption = "Character Code: Alt+" & 32 + LstChar.ListIndex
    End If
    
End Sub

Private Sub LstChar_DblClick()
    Call cmdInsert_Click
End Sub
