VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "DM Writer"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin Project1.Line3D lnRuler 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   795
      Width           =   1170
      _extentx        =   2064
      _extenty        =   53
   End
   Begin Project1.Line3D LnTop 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   795
      _extentx        =   1402
      _extenty        =   53
   End
   Begin VB.PictureBox pRuler 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   15
      Picture         =   "frmmain.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   10440
      TabIndex        =   7
      Top             =   855
      Width           =   10440
   End
   Begin RichTextLib.RichTextBox txtTmp 
      Height          =   360
      Left            =   75
      TabIndex        =   6
      Top             =   3315
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   635
      _Version        =   393217
      TextRTF         =   $"frmmain.frx":05A5
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1545
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   3060
      TabIndex        =   5
      Top             =   435
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_BOLD"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_ITALIC"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_UNDERLINE"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_STRIKE"
            Object.ToolTipText     =   "StrikeThru"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_LEFT"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_CENTER"
            Object.ToolTipText     =   "Align Center"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_RIGHT"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_BULL"
            Object.ToolTipText     =   "Bullets"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "INDENT_A"
            Object.ToolTipText     =   "Decrease Indent"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "INDENT_B"
            Object.ToolTipText     =   "Increase Indent"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_SEL"
            Object.ToolTipText     =   "Highlight"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_TEXT"
            Object.ToolTipText     =   "Text Color"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_SIZE1"
            Object.ToolTipText     =   "Increase Size"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_SIZE2"
            Object.ToolTipText     =   "Decrease Size"
            ImageIndex      =   17
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Fontsize"
      Top             =   450
      Width           =   690
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Font"
      Top             =   450
      Width           =   2205
   End
   Begin RichTextLib.RichTextBox txtEdit 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   1140
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   3413
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"frmmain.frx":0627
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14261
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_NEW"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_OPEN"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_SAVE"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_PRINT"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_CUT"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_COPY"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_PASTE"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_UNDO"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_REDO"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_FIND"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_REPLACE"
            Object.ToolTipText     =   "Replace"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_PIC"
            Object.ToolTipText     =   "Insert Picture"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_DATE"
            Object.ToolTipText     =   "Date/Time"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_SYM"
            Object.ToolTipText     =   "Symbol"
            ImageIndex      =   28
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   900
      Top             =   3300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":069E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":07B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":14F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":184A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":25D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2924
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":331A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":366C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":39BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4062
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":43B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4706
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":50FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":544E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":57A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":58B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5C04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image PicData 
      Height          =   360
      Left            =   450
      Top             =   3300
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSend 
         Caption         =   "Sen&d..."
      End
      Begin VB.Menu mnuBlank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuBlank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Cle&ar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuBlank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Goto..."
      End
      Begin VB.Menu mnublank10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace"
      End
      Begin VB.Menu mnuBlank7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &Al&l"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuBlank9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveSel 
         Caption         =   "&Save Selection"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuDateTime 
         Caption         =   "&Date and Time..."
      End
      Begin VB.Menu mnuSymbol 
         Caption         =   "&Symbol..."
      End
      Begin VB.Menu mnuPic 
         Caption         =   "&Picture..."
      End
      Begin VB.Menu mnuBlank8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile1 
         Caption         =   "&File..."
      End
   End
   Begin VB.Menu mnuPar 
      Caption         =   "Paragraph"
      Begin VB.Menu mnuSinagleLine 
         Caption         =   "&Single Line Spaceing"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "1.5 Li&ne Spaceing"
      End
      Begin VB.Menu mnuDoubleSpace 
         Caption         =   "&Double Line Spaceing"
      End
      Begin VB.Menu mnuBlank11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaLeft 
         Caption         =   "Align &Left"
      End
      Begin VB.Menu mnuaCenter 
         Caption         =   "Align &Center"
      End
      Begin VB.Menu mnuaRight 
         Caption         =   "Align &Right"
      End
      Begin VB.Menu mnuBlank12 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Begin VB.Menu mnubground 
         Caption         =   "Background..."
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About DM Writer..."
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DlgFilterIdx As Integer

Private Sub UpdateStatusbar()
    sBar1.Panels(2).Text = "Ln: " & TxtCls.LineIndex & " Col: " & TxtCls.LineLength & " Sel: " & txtEdit.SelLength
End Sub

Private Sub ClickButton(ByVal Index As Integer)
Dim mButton As Button
    Set mButton = Toolbar1.Buttons(Index)
    Call Toolbar1_ButtonClick(mButton)
    Set mButton = Nothing
End Sub

Private Sub DoFont()
On Error GoTo ErrFlag:
    With CD1
        .CancelError = True
        .Flags = (cdlCFBoth Or cdlCFApply Or cdlCFEffects)
        .FontName = cboFont.Text
        .FontSize = Val(cboSize.Text)
        
        .FontBold = Toolbar1.Buttons(1).Value
        .FontItalic = Toolbar1.Buttons(2).Value
        .FontUnderline = Toolbar1.Buttons(3).Value
        .FontStrikethru = Toolbar1.Buttons(4).Value
        'Show font dialog
        .ShowFont
        'Set Combo Box Indexs
        cboFont.ListIndex = FindInList(cboFont, .FontName)
        cboSize.ListIndex = FindInList(cboSize, .FontSize)
        'Set Font Style Buttons
        If (.FontBold) Then
            Toolbar1.Buttons(1).Value = tbrPressed
            Call ClickButton(1)
        End If
        If (.FontItalic) Then
            Toolbar1.Buttons(2).Value = tbrPressed
            Call ClickButton(2)
        End If
        If (.FontUnderline) Then
            Toolbar1.Buttons(3).Value = tbrPressed
            Call ClickButton(3)
        End If
        If (.FontStrikethru) Then
            Toolbar1.Buttons(4).Value = tbrPressed
            Call ClickButton(4)
        End If
    End With
    
    Exit Sub
ErrFlag:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Sub

Private Function GetDLGName(Optional ShowOpen As Boolean = True, Optional Title As String = "Open", Optional Filter As String)
On Error GoTo CanErr:
    'Show open or save dialog
    With CD1
        .CancelError = True
        .DialogTitle = Title
        .Filter = Filter
        
        If (ShowOpen) Then
            .ShowOpen
        Else
            .ShowSave
        End If
        
        DlgFilterIdx = .FilterIndex
        GetDLGName = .FileName
        .FileName = vbNullString
    End With
    
    Exit Function
CanErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Sub EnableMenuItems()
Dim HasSel As Boolean

    HasSel = Len(txtEdit.SelText)
    
    tBar1.Buttons(6).Enabled = HasSel
    tBar1.Buttons(7).Enabled = HasSel
    tBar1.Buttons(8).Enabled = TxtCls.CanPaste
    tBar1.Buttons(10).Enabled = TxtCls.CanUndo
    tBar1.Buttons(11).Enabled = TxtCls.CanReDo
    tBar1.Buttons(13).Enabled = Len(txtEdit.Text)
    tBar1.Buttons(14).Enabled = Len(txtEdit.Text)
    mnuCut.Enabled = HasSel
    mnuCopy.Enabled = HasSel
    mnuPaste.Enabled = TxtCls.CanPaste
    mnuClear.Enabled = HasSel
    mnuSelAll.Enabled = Len(txtEdit.Text)
    mnuUndo.Enabled = TxtCls.CanUndo
    mnuRedo.Enabled = TxtCls.CanReDo
    mnuSaveSel.Enabled = HasSel
    mnuFind.Enabled = Len(txtEdit.Text)
    mnuReplace.Enabled = Len(txtEdit.Text)
End Sub

Private Sub FontCbo(ByVal Direction As Boolean)
Dim Idx As Integer
    
    If (Direction) Then
        'Move index Down
        Idx = (cboSize.ListIndex) + 1
        If (Idx >= cboSize.ListCount) Then
            Idx = cboSize.ListIndex
        End If
    Else
        'Move Item Up
        Idx = (cboSize.ListIndex) - 1
        If (Idx <= 0) Then
            Idx = 0
        End If
    End If
    
    cboSize.ListIndex = Idx
End Sub

Private Function GetColor() As Long
On Error GoTo ErrFlag:
    
    With CD1
        .CancelError = True
        Call .ShowColor
        GetColor = .Color
    End With
    
    Exit Function
ErrFlag:
    GetColor = -1
End Function

Private Sub cboFont_Click()
On Error Resume Next
    'Set Font
    cboFont.ToolTipText = "Font: " & cboFont.Text
    txtEdit.SelFontName = cboFont.Text
    txtEdit.SetFocus
End Sub

Private Sub cboSize_Click()
On Error Resume Next
    'Set font size
    cboSize.ToolTipText = "Fontsize: " & cboSize.Text
    txtEdit.SelFontSize = Val(cboSize.Text)
    txtEdit.SetFocus
End Sub

Private Sub Command1_Click()
 MsgBox txtEdit.SelCharOffset
End Sub

Private Sub Form_Load()
Dim Cnt As Integer

    Set TxtCls = New dTxtHelper
    TxtCls.SetEditor = txtEdit

    'Add Fonts
    For Cnt = 0 To Screen.FontCount - 1
        cboFont.AddItem Screen.Fonts(Cnt)
    Next Cnt
    
    'Add Font sizes
    cboSize.AddItem "8"
    cboSize.AddItem "9"
    cboSize.AddItem "10"
    cboSize.AddItem "11"
    cboSize.AddItem "12"
    cboSize.AddItem "14"
    cboSize.AddItem "16"
    cboSize.AddItem "18"
    cboSize.AddItem "20"
    cboSize.AddItem "22"
    cboSize.AddItem "24"
    cboSize.AddItem "26"
    cboSize.AddItem "28"
    cboSize.AddItem "36"
    cboSize.AddItem "48"
    cboSize.AddItem "72"
    'Set Indexs
    cboFont.ListIndex = FindInList(cboFont, "Verdana")
    cboSize.ListIndex = 2
    'Enable/Disable menu items
    Call EnableMenuItems
    'Update statusbar
    Call UpdateStatusbar
    'Resize font combobox
    Call SetComboWidth(cboFont.Hwnd, 196)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'Resize Controls
    LnTop.Width = frmmain.ScaleWidth
    lnRuler.Width = frmmain.ScaleWidth
    txtEdit.Width = frmmain.ScaleWidth
    txtEdit.Height = (frmmain.ScaleHeight - sBar1.Height - txtEdit.Top)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Private Sub mnuAbout_Click()
    'Show about box
    Call frmAbout.Show(vbModal, frmmain)
End Sub

Private Sub mnuaCenter_Click()
    'Align Center
    Toolbar1.Buttons(7).Value = tbrPressed
    Call ClickButton(7)
End Sub

Private Sub mnuaLeft_Click()
    'Align Left
    Toolbar1.Buttons(6).Value = tbrPressed
    Call ClickButton(6)
End Sub

Private Sub mnuaRight_Click()
    'Align Right
    Toolbar1.Buttons(8).Value = tbrPressed
    Call ClickButton(8)
End Sub

Private Sub mnubground_Click()
Dim iCol As Long
    'Set editor background color
    iCol = GetColor
    If (iCol <> -1) Then
        txtEdit.BackColor = iCol
    End If
End Sub

Private Sub mnuClear_Click()
    'Clear
    Call TxtCls.Clear
    Call EnableMenuItems
End Sub

Private Sub mnuCopy_Click()
    'Copy
    Call TxtCls.Copy
    Call EnableMenuItems
End Sub

Private Sub mnuCut_Click()
    'Cut
    Call TxtCls.Cut
    Call EnableMenuItems
End Sub

Private Sub mnuDateTime_Click()
    Call frmDate.Show(vbModal, frmmain)
    If (ButtonPress = vbOK) Then
        'Insert Date and Time
        txtEdit.SelText = DateInsert
    End If
    ButtonPress = vbCancel
End Sub

Private Sub mnuDoubleSpace_Click()
    txtEdit.SelCharOffset = 160
End Sub

Private Sub mnuExit_Click()
    'Close the program
    Call Unload(frmmain)
End Sub

Private Sub mnuFile1_Click()
Dim lFile As String
    'Insert File
    lFile = GetDLGName(, "Insert File", "All Files(*.*)|*.*|")
    
    If Len(lFile) And (FindFile(lFile)) Then
        txtEdit.SelText = OpenFile(lFile)
    End If
    
End Sub

Private Sub mnuFind_Click()
    SerOp = 0 'Show only find dialog.
    frmFind.Show , frmmain
End Sub

Private Sub mnuFont_Click()
    Call DoFont
End Sub

Private Sub mnuGoto_Click()
    'Store sel position
    m_CurSelPos = txtEdit.SelStart
    'Show goto dialog
    Call frmGoto.Show(vbModal, frmmain)
    
    If (ButtonPress = vbOK) And (mGotoSelPos) Then
        txtEdit.SelStart = m_CurSelPos
    End If
    
End Sub

Private Sub mnuLine_Click()
    txtEdit.SelCharOffset = 55
End Sub

Private Sub mnuNew_Click()
    If MsgBox("Do you want to start a new document.", vbYesNo Or vbExclamation) = vbYes Then
        'Clear the editor
        txtEdit.Text = ""
    End If
End Sub

Private Sub mnuOpen_Click()
Dim lFile As String
    'Open new file
    lFile = GetDLGName(, , "RTF Files(*.rtf)|*.rtf|")
    'Check name and if file is found
    If Len(lFile) And FindFile(lFile) Then
        Call txtEdit.LoadFile(lFile, 0)
    End If
    
End Sub

Private Sub mnuPageSetup_Click()
On Error GoTo PrnErr:
    ' Show printer dialog
    With CD1
        .CancelError = True
        .DialogTitle = "Page Setup"
        .ShowPrinter
    End With
    
    Exit Sub
PrnErr:
    If (Err.Number <> cdlCancel) Then
        MsgBox Err.Description, vbCritical, "Error#" & Err.Number
    End If
End Sub

Private Sub mnuPaste_Click()
    'Paste
    Call TxtCls.Paste
    Call EnableMenuItems
End Sub

Private Sub mnuPic_Click()
Dim lFile As String
    lFile = GetDLGName(, "Picture", "Picture Files(*.bmp)|*.bmp|GIF Files(*.gif)|*.gif|JPEG Files(*.jpg)|*.jpg|")
    
    If Len(lFile) Then
        PicData.Picture = LoadPicture(lFile)
        'Copy picture to clipboard
        Call Clipboard.SetData(PicData.Picture, vbCFBitmap)
        'Paste in the picture
        Call TxtCls.Paste
        'Destroy the picture
        Set PicData.Picture = Nothing
    End If
End Sub

Private Sub mnuPrint_Click()
On Error GoTo PrnErr:

        With CD1
            .CancelError = True
            .DialogTitle = "Print"
            .Flags = (cdlPDReturnDC Or cdlPDNoPageNums)
            
            If (txtEdit.SelLength = 0) Then
                .Flags = (.Flags Or cdlPDAllPages)
            Else
                .Flags = (.Flags Or cdlPDSelection)
            End If
            'Show print dialog.
            .ShowPrinter
            'Print document.
            Call txtEdit.SelPrint(.hDC)
        End With
    
    Exit Sub
PrnErr:
    If (Err.Number <> cdlCancel) Then
        MsgBox Err.Description, vbCritical, "Error#" & Err.Number
    End If

End Sub

Private Sub mnuRedo_Click()
    'Redo
    Call TxtCls.Undo
    tBar1.Buttons(10).Enabled = True
    tBar1.Buttons(11).Enabled = False
    mnuRedo.Enabled = False
    mnuUndo.Enabled = True
End Sub

Private Sub mnuReplace_Click()
    SerOp = 1 'Shows find and replace dialog.
    frmFind.Show , frmmain
End Sub

Private Sub mnuSave_Click()
    Call mnuSaveAs_Click
End Sub

Private Sub mnuSaveAs_Click()
Dim lFile As String
    'Save file
    lFile = GetDLGName(False, "Save As", "RTF Files(*.rtf)|*.rtf|Text Files(*.txt)|*.txt|")

    If Len(lFile) Then
        If (DlgFilterIdx = 1) Then
            'Save as RTF
            Call txtEdit.SaveFile(lFile, 0)
        Else
            'Save as normal Text
            Call txtEdit.SaveFile(lFile, 1)
        End If
    End If
End Sub

Private Sub mnuSaveSel_Click()
Dim lFile As String
    
    'Save selection
    lFile = GetDLGName(False, "Save Selection", "RTF Files(*.rtf)|*.rtf|Text Files(*.txt)|*.txt|")
    'Check for file name
    If Len(lFile) Then
        If (DlgFilterIdx = 1) Then
            'Save as RTF
            txtTmp.TextRTF = txtEdit.SelRTF
            Call txtTmp.SaveFile(lFile, 0)
        Else
            'Save as normal Text
            Call txtEdit.SaveFile(lFile, 1)
        End If
    End If
    
End Sub

Private Sub mnuSelAll_Click()
    'Select All
    Call TxtCls.SelectAll
    Call EnableMenuItems
End Sub

Private Sub mnuSend_Click()
    'Start up Email
    Call RunApp(frmmain.Hwnd, "open", "mailto:yourname@yourmail.com?subject=Subject&body=" & txtEdit.Text)
End Sub

Private Sub mnuSinagleLine_Click()
    'Single line spaceing
    txtEdit.SelCharOffset = 0
End Sub

Private Sub mnuSymbol_Click()
    'Insert Symbol
    Call frmChar.Show(vbModal, frmmain)
    If (ButtonPress = vbOK) Then
        txtEdit.SelFontName = SelFont
        txtEdit.SelText = CharInsert
        txtEdit.SelFontName = cboFont.Text
    End If
    ButtonPress = vbCancel
End Sub

Private Sub mnuUndo_Click()
    'Undo
    Call TxtCls.Undo

    tBar1.Buttons(10).Enabled = False
    tBar1.Buttons(11).Enabled = True
    mnuRedo.Enabled = True
    mnuUndo.Enabled = False

End Sub

Private Sub tBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "M_NEW"
            Call mnuNew_Click
        Case "M_OPEN"
            Call mnuOpen_Click
        Case "M_SAVE"
            Call mnuSaveAs_Click
        Case "M_PRINT"
            Call mnuPrint_Click
        Case "M_CUT"
            Call mnuCut_Click
        Case "M_COPY"
            Call mnuCopy_Click
        Case "M_PASTE"
            Call mnuPaste_Click
        Case "M_UNDO"
            Call mnuUndo_Click
        Case "M_REDO"
            Call mnuRedo_Click
        Case "M_FIND"
            Call mnuFind_Click
        Case "M_REPLACE"
            Call mnuReplace_Click
        Case "M_PIC"
            Call mnuPic_Click
        Case "M_DATE"
            Call mnuDateTime_Click
        Case "M_SYM"
            Call mnuSymbol_Click
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim TmpColor As Long

    Select Case Button.Key

        Case "M_BOLD"
            txtEdit.SelBold = Button.Value
        Case "M_ITALIC"
            txtEdit.SelItalic = Button.Value
        Case "M_UNDERLINE"
            txtEdit.SelUnderline = Button.Value
        Case "M_STRIKE"
            txtEdit.SelStrikeThru = Button.Value
        Case "M_LEFT", "M_CENTER", "M_RIGHT"
            Toolbar1.Buttons(6).Value = tbrUnpressed
            Toolbar1.Buttons(7).Value = tbrUnpressed
            Toolbar1.Buttons(8).Value = tbrUnpressed
            Toolbar1.Buttons(Button.Index).Value = tbrPressed
            'Check what button was pressed and set sel alignment
            If (Button.Index = 6) Then txtEdit.SelAlignment = 0
            If (Button.Index = 7) Then txtEdit.SelAlignment = 2
            If (Button.Index = 8) Then txtEdit.SelAlignment = 1
        Case "M_BULL"
            'Dispay Bullets
            txtEdit.SelBullet = Button.Value
        Case "INDENT_A"
            txtEdit.SelIndent = (txtEdit.SelIndent - 400)
            Call txtEdit.SetFocus
        Case "INDENT_B"
            txtEdit.SelIndent = (txtEdit.SelIndent + 400)
            Call txtEdit.SetFocus
        Case "M_SEL"
            TmpColor = GetColor
            If (TmpColor <> -1) Then
                'Set the text sel color
                Call HighLight(txtEdit, TmpColor)
            End If
        Case "M_TEXT"
            TmpColor = GetColor
            If (TmpColor <> -1) Then
                'Set the text color
                txtEdit.SelColor = TmpColor
            End If
        Case "M_SIZE1"
            Call FontCbo(True)
        Case "M_SIZE2"
            Call FontCbo(False)
    End Select
    
End Sub

Private Sub txtEdit_Click()
    Call UpdateStatusbar
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    'Check if paste key is used.
    If (KeyCode = vbKeyV) And (Shift = 2) Then
        Call mnuPaste_Click
        'Set keycode to zero
        KeyCode = 0
    End If
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UpdateStatusbar
End Sub

Private Sub txtEdit_SelChange()
On Error Resume Next
    SelFindStr = txtEdit.SelText
    'Enable/Disable menu items
    Call EnableMenuItems
    Call UpdateStatusbar
    'Update font styles
    Toolbar1.Buttons(1).Value = Abs(txtEdit.SelBold)
    Toolbar1.Buttons(2).Value = Abs(txtEdit.SelItalic)
    Toolbar1.Buttons(3).Value = Abs(txtEdit.SelUnderline)
    Toolbar1.Buttons(4).Value = Abs(txtEdit.SelStrikeThru)
End Sub

