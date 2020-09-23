VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dTxtHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Const WM_CUT As Long = &H300
Private Const WM_COPY As Long = &H301
Private Const WM_PASTE As Long = &H302
Private Const WM_CLEAR As Long = &H303
Private Const WM_UNDO As Long = &H304
Private Const WM_USER As Long = &H400

Private Const EM_GETFIRSTVISIBLELINE As Long = &HCE
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_GETLINE As Long = &HC4
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_CHARFROMPOS As Long = &HD7
Private Const EM_LINELENGTH As Long = &HC1
Private Const EM_CANUNDO As Long = &HC6
Private Const EM_CANREDO As Long = (WM_USER + 85)
Private Const EM_CANPASTE As Long = (WM_USER + 50)
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_SETTARGETDEVICE As Long = (WM_USER + 72)
Private Const EM_SETTABSTOPS As Long = &HCB

Private Const GWL_STYLE As Long = -16
Private Const WS_HSCROLL As Long = &H100000
Private Const ES_AUTOHSCROLL As Long = &H80&

Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4

Enum ConvertOP
    dUpperCase = 0
    dLowerCase = 1
    dTitleCase = 2
    dInvertCase = 3
    dROT13 = 4
End Enum

Private m_GotoLine As Long
Private WithEvents Editor As RichTextBox
Attribute Editor.VB_VarHelpID = -1

Public Sub WordWrap(Wrap As Boolean)
Dim OrgWnd As Long
    
    'Get the orignal window style
    OrgWnd = GetWindowLong(Editor.Hwnd, GWL_STYLE)
    
    If (Wrap) Then
        OrgWnd = OrgWnd Or (WS_HSCROLL Or ES_AUTOHSCROLL)
        SendMessage Editor.Hwnd, EM_SETTARGETDEVICE, 0, ByVal 0&
    Else
        OrgWnd = OrgWnd And Not (WS_HSCROLL Or ES_AUTOHSCROLL)
        SendMessage Editor.Hwnd, EM_SETTARGETDEVICE, 0, ByVal 1&
    End If
    
    'Set the window style
    SetWindowLong Editor.Hwnd, GWL_STYLE, OrgWnd
    SetWindowPos Editor.Hwnd, 0, 0, 0, 0, 0, (SWP_DRAWFRAME Or SWP_FRAMECHANGED Or SWP_NOMOVE _
    Or SWP_NOOWNERZORDER Or SWP_NOSIZE Or SWP_NOZORDER)

End Sub

Public Sub Convert(ConvOp As ConvertOP)
Dim iStart As Long
Dim iLen As Long
    
    iStart = Editor.SelStart
    iLen = Editor.SelLength
    
    Select Case ConvOp
        Case dUpperCase
            'Convert to uppercase
            If Len(Editor.SelText) <> 0 Then
                Editor.SelText = UCase(Editor.SelText)
            Else
                Editor.Text = UCase(Editor.Text)
            End If
        Case dLowerCase
            'Convert to lowercase
            If Len(Editor.SelText) <> 0 Then
                Editor.SelText = LCase(Editor.SelText)
            Else
                Editor.Text = LCase(Editor.Text)
            End If
        Case dTitleCase
            'Convert to title case
            If Len(Editor.SelText) <> 0 Then
                Editor.SelText = StrConv(Editor.SelText, vbProperCase)
            Else
                Editor.Text = StrConv(Editor.Text, vbProperCase)
            End If
        Case dInvertCase
            'Invert text case
            If Len(Editor.SelText) <> 0 Then
                Editor.SelText = InvertCase(Editor.SelText)
            Else
                Editor.Text = InvertCase(Editor.Text)
            End If
    End Select
    
    If (iLen > 0) Then
        Call SelectText(iStart, iLen)
    End If
    
End Sub

Public Sub Cut()
    SendMessage Editor.Hwnd, WM_CUT, 0, 0
End Sub

Public Sub Clear()
    SendMessage Editor.Hwnd, WM_CLEAR, 0, 0
End Sub

Public Sub Copy()
    SendMessage Editor.Hwnd, WM_COPY, 0, 0
End Sub

Public Sub Undo()
    SendMessage Editor.Hwnd, WM_UNDO, 0, 0
End Sub

Public Sub Paste()
    SendMessage Editor.Hwnd, WM_PASTE, 0, 0
End Sub

Public Sub SelectAll()
    Editor.SelStart = 0
    Editor.SelLength = Len(Editor.Text)
    Editor.SetFocus
End Sub

Public Property Get SelPos() As Long
    SelPos = Editor.SelStart
End Property

Public Property Let SelPos(ByVal vNewSelPos As Long)
    Editor.SelStart = vNewSelPos
End Property

Public Sub SelectText(ByVal Start As Long, ByVal Length As Long)
    Editor.SelStart = Start
    Editor.SelLength = Length
    Editor.SetFocus
End Sub

Public Property Let SetEditor(ByVal vNewValue As RichTextBox)
    Set Editor = vNewValue
End Property

Public Property Get CanPaste() As Boolean
    CanPaste = Len(Clipboard.GetText(vbCFText)) <> False
End Property

Public Property Get LineCount() As Long
    LineCount = SendMessage(Editor.Hwnd, EM_GETLINECOUNT, 0, 0)
End Property

Public Property Get VisableLine() As Long
    VisableLine = SendMessage(Editor.Hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)
End Property

Public Property Get LineIndex() As Long
    LineIndex = SendMessage(Editor.Hwnd, EM_LINEFROMCHAR, (Editor.SelStart + Editor.SelLength), 0) + 1
End Property

Public Property Get LineLength() As Variant
    LineLength = SendMessage(Editor.Hwnd, EM_LINELENGTH, Editor.SelStart, 0) + 1
End Property

Public Sub LoadFromFile(ByVal FileName As String)
Dim fp As Long
Dim Buff As String
Dim Bytes() As Byte

    fp = FreeFile
    
    Open FileName For Binary As #fp
        If LOF(fp) > 0 Then
            ReDim Bytes(0 To LOF(fp) - 1)
            Get #fp, , Bytes
        End If
    Close #fp
    
    Buff = StrConv(Bytes, vbUnicode)
    'Check for unicode
    If Left(Buff, 2) = "ÿþ" Or Left$(Buff, 2) = "þÿ" Then
        Buff = Replace(Right$(Buff, Len(Buff) - 2), Chr$(0), "")
    End If
    
    Editor.Text = Buff
    Buff = vbNullString
    Erase Bytes
End Sub

Public Sub SaveToFile(ByVal FileName As String)
Dim fp As Long
    fp = FreeFile
    Open FileName For Output As #fp
        Print #fp, Editor.Text;
    Close #fp
End Sub

Public Sub SaveSelection(ByVal FileName As String)
Dim fp As Long
    fp = FreeFile
    Open FileName For Output As #fp
        Print #fp, Editor.SelText;
    Close #fp
End Sub

Public Property Get CanUndo() As Boolean
    CanUndo = SendMessage(Editor.Hwnd, EM_CANUNDO, 0, 0)
End Property

Public Property Get CanReDo() As Boolean
    CanReDo = SendMessage(Editor.Hwnd, EM_CANREDO, 0, 0)
End Property

Public Property Get GotoLine() As Long
    GotoLine = m_GotoLine
End Property

Public Property Let GotoLine(ByVal vNewLine As Long)
Dim Ret As Long
On Error Resume Next

    m_GotoLine = vNewLine
    Ret = SendMessage(Editor.Hwnd, EM_LINEINDEX, (m_GotoLine - 1), 0)
    
    With Editor
        .SelStart = Ret
        .SelLength = 0
        .SetFocus
    End With
    
End Property

Function GetLineText(ByVal Index As Long)
Dim sBuff As String
Dim Ret As Long

    sBuff = Space(255)
    Ret = SendMessage(Editor.Hwnd, EM_GETLINE, Index - 1, ByVal sBuff)
    GetLineText = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    
    sBuff = vbNullString
    
End Function

Private Function InvertCase(StrSrc As String) As String
Dim Bytes() As Byte
Dim c As Byte
Dim ch As Byte
Dim Count As Long
    
    'Convert to byte array
    Bytes = StrConv(StrSrc, vbFromUnicode)
    'Check we have data to work with
    If UBound(Bytes) < 0 Then
        Exit Function
    End If
    
    For Count = 0 To UBound(Bytes)
        'Get char
        c = Bytes(Count)
        ch = Asc(UCase(Chr(c)))
        'Check if char code is upper
        If (c = ch) Then
            'Set new chat code to lowercase
            c = Asc(LCase(Chr(c)))
        Else
            'Set char code to lowercase
            c = Asc(UCase(Chr(c)))
        End If
        Bytes(Count) = c
    Next Count
    
    'Ret string
    InvertCase = StrConv(Bytes, vbUnicode)
    'Clear up
    Erase Bytes
    c = 0
    ch = 0
    Count = 0
End Function

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
On Error Resume Next
Dim oStart As Integer
    Editor.HideSelection = True
    Editor.SelStart = 0
    Editor.SelLength = Len(Editor.Text)
    Editor.SelColor = NewColor
    Editor.SelStart = 0
    Editor.HideSelection = False
End Property

Public Sub TabStops(TabStopSize As Long)
Dim mTab(1) As Long
    mTab(0) = TabStopSize
    SendMessage Editor.Hwnd, EM_SETTABSTOPS, 4, mTab(0)
End Sub

