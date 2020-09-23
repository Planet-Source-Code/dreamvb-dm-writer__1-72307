VERSION 5.00
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date and Time"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3165
      TabIndex        =   3
      Top             =   1095
      Width           =   1230
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   390
      Left            =   3165
      TabIndex        =   2
      Top             =   630
      Width           =   1230
   End
   Begin VB.ListBox LstFormat 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   630
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a format:"
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   300
      Width           =   1110
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Call Unload(frmDate)
End Sub

Private Sub cmdInsert_Click()
    ButtonPress = vbOK
    Call Unload(frmDate)
End Sub

Private Sub Form_Load()
    'Add date and time formats
    LstFormat.AddItem Format(Date, "DD/MM/YYYY")
    LstFormat.AddItem Format(Date, "DD/MM/YY")
    LstFormat.AddItem Format(Date, "DD/M/YY")
    LstFormat.AddItem Format(Date, "DD.M.YY")
    LstFormat.AddItem Format(Date, "YYYY-MM-DD")
    LstFormat.AddItem Format(Date, "M/DD/YYYY")
    LstFormat.AddItem Format(Date, "M/DD/YY")
    LstFormat.AddItem Format(Date, "MM/DD/YY")
    LstFormat.AddItem Format(Date, "MM/DD/YYYY")
    LstFormat.AddItem Format(Date, "DD MMMM YYYY")
    LstFormat.AddItem Format(Date, "DDDD, MMMM DD, YYYY")
    
    LstFormat.AddItem Format(Time, "HH:MM:SS AMPM")
    LstFormat.AddItem Format(Time, "HH:MM AMPM")
    LstFormat.AddItem Format(Time, "HH:MM:SS")
    LstFormat.AddItem Format(Time, "HH:MM")
    
    Set frmDate.Icon = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDate = Nothing
End Sub

Private Sub LstFormat_Click()
    DateInsert = LstFormat.Text
End Sub

Private Sub LstFormat_DblClick()
    Call cmdInsert_Click
End Sub
