Attribute VB_Name = "Module1"
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = 0
Public Const NIM_MODIFY = 1
Public Const NIM_DELETE = 2
Public Const NIF_MESSAGE = 1
Public Const NIF_ICON = 2
Public Const NIF_TIP = 4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONDBLCLK = &H206
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const ABM_REMOVE = &H1
Public Const ABM_GETSTATE = &H4
Public Const ABM_SETAUTOHIDEBAR = &H8

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Global picWidth As Integer, picHeight As Integer
Global posX As Integer, posY As Integer
Global firsttime  As Boolean
Global bg, x, y, buff, pback
Global ind As Integer

Public Sub AddSysTrayIcon(ByRef frm As Form, ByVal strToolTip As String)
    
    Shell_NotifyIconA NIM_ADD, setNOTIFYICONDATA( _
        hwnd:=frm.hwnd, _
        ID:=vbNull, _
        Flags:=NIF_MESSAGE Or NIF_ICON Or NIF_TIP, _
        CallbackMessage:=WM_MOUSEMOVE, _
        Icon:=frm.Icon, _
        Tip:=strToolTip)
    
End Sub

Public Sub KillSysTrayIcon(ByRef frm As Form)

    Shell_NotifyIconA NIM_DELETE, setNOTIFYICONDATA( _
        hwnd:=frm.hwnd, _
        ID:=vbNull, _
        Flags:=NIF_MESSAGE Or NIF_ICON Or NIF_TIP, _
        CallbackMessage:=WM_MOUSEMOVE, _
        Icon:=frm.Icon, _
        Tip:="")

End Sub

Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, Flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA

    Dim nidTemp As NOTIFYICONDATA
   
    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = ID
    nidTemp.uFlags = Flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr(0)

    setNOTIFYICONDATA = nidTemp
    
End Function

Function cycle()

If Not firsttime Then
    ' draw the background back from the buff(er). buffer->desktop
    rc% = BitBlt(bg, posX, posY, picWidth, picHeight, buff, 0, 0, vbSrcCopy)
    ' move co-ords.
    posX = posX + 1.3
    If Int(posX) = Screen.Width / 15 Then       ' hits edge of screen, set it back.
        posX = -86
    End If
    posY = posY - 2.5
    If Int(posY) = -86 Then
        posY = Screen.Height / 15
    End If
End If

ind = ind + 1
If ind > 3 Then
    ind = 0
End If

'save new background to buff(er). desktop->buffer
rc% = BitBlt(buff, 0, 0, picWidth, picHeight, bg, posX, posY, vbSrcCopy)
' create 'sprite'.
rc% = BitBlt(pback, x, y, picWidth, picHeight, buff, 0, 0, vbSrcCopy) 'masks
rc% = BitBlt(pback, x, y, picWidth, picHeight, Form1.mask(ind).hdc, 0, 0, vbSrcPaint) 'masks
rc% = BitBlt(pback, x, y, picWidth, picHeight, Form1.pic(ind).hdc, 0, 0, vbSrcAnd) 'masks

'draw sprite. sprite->desktop
rc% = BitBlt(bg, posX, posY, picWidth, picHeight, pback, 0, 0, vbSrcCopy)  'copies product To screen.
firsttime = False

ReleaseDC hwnddesk, hdcdesk

End Function


