VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Goto"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstGoto 
      Height          =   840
      Left            =   150
      TabIndex        =   5
      Top             =   390
      Width           =   2010
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3870
      TabIndex        =   2
      Top             =   915
      Width           =   1110
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2625
      TabIndex        =   1
      Top             =   915
      Width           =   1110
   End
   Begin VB.TextBox txtGoto 
      Height          =   350
      Left            =   2595
      TabIndex        =   0
      Text            =   "1"
      Top             =   390
      Width           =   2370
   End
   Begin VB.Label lblGoto1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GoTo:"
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   450
   End
   Begin VB.Label lblGoto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line Number:"
      Height          =   195
      Left            =   2625
      TabIndex        =   3
      Top             =   150
      Width           =   945
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    ButtonPress = vbCancel
    Unload frmGoto
End Sub

Private Sub cmdOK_Click()
    mGotoSelPos = False
    TxtCls.GotoLine = Val(txtGoto.Text)

    Select Case lstGoto.ListIndex
        Case 0
            'Start of document
            TxtCls.GotoLine = 1
        Case 1
            'End of document
            TxtCls.GotoLine = TxtCls.LineCount
        Case 2
            'Line number
            If (TxtCls.GotoLine = 0) Then
                TxtCls.GotoLine = 1
            End If
        Case 3
            'Position
            mGotoSelPos = True
            m_CurSelPos = Val(txtGoto.Text)
    End Select
    
    ButtonPress = vbOK
    Unload frmGoto
End Sub

Private Sub Form_Activate()
    txtGoto.SelLength = Len(txtGoto.Text)
    txtGoto.SetFocus
End Sub

Private Sub Form_Load()
    Set frmGoto.Icon = Nothing
    lstGoto.AddItem "Start Of Document"
    lstGoto.AddItem "End of Document"
    lstGoto.AddItem "Line Number"
    lstGoto.AddItem "Position"
    lstGoto.ListIndex = 2
    txtGoto.Text = TxtCls.GotoLine
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGoto = Nothing
End Sub

Private Sub lstGoto_Click()
On Error Resume Next

    txtGoto.Enabled = (lstGoto.ListIndex = 2) Or (lstGoto.ListIndex = 3)
    
    If (lstGoto.ListIndex = 2) Then
        txtGoto.SelLength = Len(txtGoto.Text)
        txtGoto.SetFocus
    End If
    
    If (lstGoto.ListIndex = 3) Then
       txtGoto.Text = m_CurSelPos
    End If
    
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case 13
            KeyAscii = 0
            Call cmdOK_Click
        Case Else
            KeyAscii = 0
    End Select
End Sub
