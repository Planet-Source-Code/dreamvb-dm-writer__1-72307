VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4215
      TabIndex        =   6
      Top             =   1410
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4215
      TabIndex        =   5
      Top             =   975
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox txtReplace 
      Height          =   330
      Left            =   1155
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.CheckBox chkMatch 
      Caption         =   "Match &case"
      Height          =   225
      Left            =   90
      TabIndex        =   2
      Top             =   1560
      Width           =   1635
   End
   Begin VB.TextBox txtFind 
      Height          =   330
      Left            =   1155
      TabIndex        =   0
      Top             =   165
      Width           =   2790
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4215
      TabIndex        =   4
      Top             =   555
      Width           =   1110
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Left            =   4215
      TabIndex        =   3
      Top             =   150
      Width           =   1110
   End
   Begin VB.Label lblReplace 
      AutoSize        =   -1  'True
      Caption         =   "Replace with:"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   660
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "Find what:"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   225
      Width           =   735
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sPos As Integer
Private Found As Boolean

Private Sub cmdCancel_Click()
    Call Unload(frmFind)
End Sub

Private Sub cmdFindNext_Click()
Dim sComp As VbCompareMethod

    If (chkMatch) Then
        sComp = vbBinaryCompare
    Else
        sComp = vbTextCompare
    End If

    sPos = InStr(sPos + 1, frmmain.txtEdit.Text, txtFind.Text, sComp)
    
    If (sPos > 0) Then
        Call TxtCls.SelectText(sPos - 1, Len(txtFind.Text))
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
    Else
        MsgBox "String not found '" & txtFind.Text & "'", vbInformation, "Find"
        'Disable command buttons
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    End If
    
End Sub

Private Sub cmdReplace_Click()
    If Len(frmmain.txtEdit.SelText) > 0 Then
        frmmain.txtEdit.SelStart = (sPos - 1)
        frmmain.txtEdit.SelLength = Len(txtFind.Text)
        frmmain.txtEdit.SelText = txtReplace.Text
    End If
    
    Call cmdFindNext_Click
    
End Sub

Private Sub cmdReplaceAll_Click()
    Do Until (sPos = 0)
        Call cmdReplace_Click
        DoEvents
    Loop
End Sub

Private Sub Form_Load()
On Error Resume Next
    'Destroy icon
    Set frmFind.Icon = Nothing
    
    txtFind.Text = SelFindStr
    sPos = 1
    
    If (SerOp = 0) Then
        frmFind.Caption = "Find"
    End If
    
    If (SerOp = 1) Then
        frmFind.Caption = "Replace"
        lblReplace.Visible = True
        txtReplace.Visible = True
        cmdReplace.Visible = True
        cmdReplaceAll.Visible = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFind = Nothing
End Sub

Private Sub txtFind_Change()
    cmdFindNext.Enabled = Len(txtFind.Text)
End Sub
