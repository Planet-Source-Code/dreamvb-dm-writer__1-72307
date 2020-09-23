Attribute VB_Name = "Tools"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
'Variables
Public ButtonPress As VbMsgBoxResult
Public DateInsert As String
Public CharInsert As String
Public PasteFormat As Integer
'Goto Variables
Public m_CurSelPos As Long
Public mGotoSelPos As Boolean
'Find Variables
Public SerOp As Integer
Public SelFindStr As String

Public SelFont As String
Public TxtCls As dTxtHelper
'Consts
Private Const CB_SETDROPPEDWIDTH As Long = &H160

Public Sub SetComboWidth(ByVal Hwnd As Long, ByVal Width As Long)
Dim iRet As Long
    iRet = SendMessage(Hwnd, CB_SETDROPPEDWIDTH, Width, 0)
End Sub

Private Function AddColorToTable(strRTF As String, lColor As Long) As Integer
Dim iPos As Long, jpos As Long
'I never wote this code found on the internet but like to says thanks to who ever did do it
Dim ctbl As String
Dim tagColors
Dim nColors As Integer
Dim tagNew As String
Dim i As Integer
Dim iLen As Integer
Dim split1 As String
Dim split2 As String

    'make new color into tag
    tagNew = "\red" & CStr(lColor And &HFF) & _
        "\green" & CStr(Int(lColor / &H100) And &HFF) & _
        "\blue" & CStr(Int(lColor / &H10000))
    
    'find colortable
    iPos = InStr(strRTF, "{\colortbl")
    
    If iPos > 0 Then 'if table already exists
        jpos = InStr(iPos, strRTF, ";}")
        'color table
        ctbl = Mid(strRTF, iPos + 12, jpos - iPos - 12)
        'array of color tags
        tagColors = Split(ctbl, ";")
        nColors = UBound(tagColors) + 2
        'see if our color already exists in table
        For i = 0 To UBound(tagColors)
            If tagColors(i) = tagNew Then
                AddColorToTable = i + 1
                Exit Function
            End If
        Next i
        
        split1 = Left(strRTF, jpos)
        split2 = Mid(strRTF, jpos + 1)
        strRTF = split1 & tagNew & ";" & split2
        AddColorToTable = nColors
    
    Else
        'color table doesn't exists, let's make one
        iPos = InStr(strRTF, "{\fonttbl") 'beginning of font table
        jpos = InStr(iPos, strRTF, ";}}") + 2 'end of font table
        split1 = Left(strRTF, jpos)
        split2 = Mid(strRTF, jpos + 1)
        strRTF = split1 & "{\colortbl ;" & tagNew & ";}" & split2
        AddColorToTable = 1
    End If

End Function

Public Sub HighLight(RTB As RichTextBox, lColor As Long)
'I never wote this code found on the internet but like to says thanks to who ever did do it
Dim iPos As Long
Dim strRTF As String
Dim bkColor As Integer
    With RTB
        iPos = .SelStart
        'bracket selection
        .SelText = Chr(&H9D) & .SelText & Chr(&H81)
        strRTF = RTB.TextRTF
        'add new color
        bkColor = AddColorToTable(strRTF, lColor)
        'add highlighting
        strRTF = Replace(strRTF, "\'9d", "\up1\highlight" & CStr(bkColor) & "")
        strRTF = Replace(strRTF, "\'81", "\highlight0\up0 ")

        .TextRTF = strRTF
        .SelStart = iPos
       End With
End Sub

Public Sub RunApp(iHwnd As Long, OpenOp As String, FileName As String)
Dim Ret As Long
    Ret = ShellExecute(iHwnd, OpenOp, FileName, "", "", 1)
End Sub

Public Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Public Function FindInList(Cbo As ComboBox, StrFind As String) As Integer
Dim x As Integer
Dim Idx As Integer
    'Locate an items index in a combobox
    For x = 0 To Cbo.ListCount
        If LCase(StrFind) = LCase(Cbo.List(x)) Then
            Idx = x
            Exit For
        End If
    Next x
    
    FindInList = Idx
End Function

Public Function OpenFile(FileName As String) As String
Dim fp As Long
Dim Bytes() As Byte
    
    fp = FreeFile
    
    Open FileName For Binary As #fp
        If LOF(fp) > 0 Then
            ReDim Bytes(0 To LOF(fp) - 1)
            Get #fp, , Bytes
        End If
    Close #fp
    
    OpenFile = StrConv(Bytes, vbUnicode)
    Erase Bytes
End Function

