VERSION 5.00
Begin VB.Form frmPasteS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paste Special"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3705
      TabIndex        =   3
      Top             =   1260
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   3705
      TabIndex        =   2
      Top             =   720
      Width           =   1230
   End
   Begin VB.ListBox lstPaste 
      Height          =   1500
      IntegralHeight  =   0   'False
      Left            =   270
      TabIndex        =   0
      Top             =   720
      Width           =   3315
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paste As:"
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   465
      Width           =   675
   End
End
Attribute VB_Name = "frmPasteS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Call Unload(frmPasteS)
End Sub

Private Sub cmdOK_Click()
    ButtonPress = vbOK
    Call Unload(frmPasteS)
End Sub

Private Sub Form_Load()
    'Remove icon
    Set frmPasteS.Icon = Nothing
    lstPaste.AddItem "Text"
    lstPaste.AddItem "Bitmap"
    lstPaste.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPasteS = Nothing
End Sub

Private Sub lstPaste_Click()
    'Store the index
    PasteFormat = lstPaste.ListIndex
End Sub

Private Sub lstPaste_DblClick()
    Call cmdOK_Click
End Sub
