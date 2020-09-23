VERSION 5.00
Begin VB.UserControl tlghnUC96315 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   254
   Begin VB.Menu mnuTolgahan 
      Caption         =   "tolgahan"
      Begin VB.Menu mnuCalistir 
         Caption         =   "Çalýþtýr"
      End
      Begin VB.Menu mnuSil 
         Caption         =   "Sil"
      End
      Begin VB.Menu mnuCizgi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOzellikler 
         Caption         =   "Özellikler"
      End
   End
End
Attribute VB_Name = "tlghnUC96315"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const BDR_INNER = &HC

Private Const BDR_OUTER = &H3

Private Const BDR_RAISED = &H5

Private Const BDR_RAISEDINNER = &H4

Private Const BDR_RAISEDOUTER = &H1

Private Const BDR_SUNKEN = &HA

Private Const BDR_SUNKENINNER = &H8

Private Const BDR_SUNKENOUTER = &H2

Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)

Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Const BF_BOTTOM = &H8

Private Const BF_LEFT = &H1

Private Const BF_RIGHT = &H4

Private Const BF_TOP = &H2

Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Const DT_CENTER = &H1
Private Const DT_WORDBREAK = &H10

Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Const MAX_PATH = 260

Private Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: type name
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Default Property Values:
Const m_def_OtomatikCalistir = True
Const m_def_Dosya = "C:\WINDOWS\Explorer.exe"
'Property Variables:
Dim m_OtomatikCalistir As Boolean
Dim m_Dosya As String
Dim hIcon As Long, isim As String
Private Const SHGFI_ICON = &H100                         '  get icon
'Event Declarations:
Event Tikla() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Tikla.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

Private Sub tolgahanKenarlikCiz(Optional durum As Byte)
    Cls
    Dim rt As RECT
    SetRect rt, 0, 0, ScaleWidth, ScaleHeight
    Select Case durum
        Case 0
            DrawEdge hDC, rt, EDGE_ETCHED, BF_RECT
        Case 1
            DrawEdge hDC, rt, EDGE_RAISED, BF_RECT
        Case 2
            DrawEdge hDC, rt, EDGE_SUNKEN, BF_RECT
    End Select
End Sub
Private Sub tolgahanBaslikCiz(Optional durum As Byte)
    Dim rt As RECT
    SetRect rt, 0, ScaleHeight - TextHeight("A") * 2.1, ScaleWidth, ScaleHeight
    Select Case durum
        Case 1: OffsetRect rt, -1, -1
        Case 2: OffsetRect rt, 1, 1
    End Select
    DrawTextEx hDC, isim, Len(isim), rt, DT_CENTER Or DT_WORDBREAK, ByVal 0&
End Sub
Private Sub tolgahanSimgeCiz(Optional durum As Byte)
    Dim rt As RECT
    SetRect rt, (ScaleWidth - 32) / 2, (ScaleHeight - TextHeight("A") * 2.1 - 32) / 2, 0, 0
    Select Case durum
        Case 1: OffsetRect rt, -1, -1
        Case 2: OffsetRect rt, 1, 1
    End Select
    DrawIconEx hDC, rt.Left, rt.Top, hIcon, 32, 32, 0, 0, 3
End Sub
Private Sub tolgahanDugmeCiz(Optional durum As Byte)
    tolgahanKenarlikCiz durum
    tolgahanBaslikCiz durum
    tolgahanSimgeCiz durum
End Sub
Public Sub calistir()
    ShellExecute 0, "", m_Dosya, "", "", 1
End Sub
Private Sub UserControl_DblClick()
tolgahanDugmeCiz 2
UserControl_Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    tolgahanDugmeCiz 2
Else
    PopupMenu mnuTolgahan, , , , mnuCalistir
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 0 Then Exit Sub
If (x > 2 And x < ScaleWidth - 2 And y > 2 And y < ScaleHeight - 2) Then
    tolgahanDugmeCiz 1
    SetCapture hwnd
Else
    ReleaseCapture
    tolgahanDugmeCiz
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    tolgahanDugmeCiz
End Sub

Private Sub UserControl_Resize()
    If ScaleWidth < 120 Then Width = 1800
    If ScaleHeight < 90 Then Height = 1350
    tolgahanDugmeCiz
End Sub

Private Sub UserControl_Show()
    tolgahanDosyaBilgiAl
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_OtomatikCalistir = m_def_OtomatikCalistir
    m_Dosya = m_def_Dosya
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_OtomatikCalistir = PropBag.ReadProperty("OtomatikCalistir", m_def_OtomatikCalistir)
    m_Dosya = PropBag.ReadProperty("Dosya", m_def_Dosya)
End Sub

Private Sub UserControl_Terminate()
    DestroyIcon hIcon
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("OtomatikCalistir", m_OtomatikCalistir, m_def_OtomatikCalistir)
    Call PropBag.WriteProperty("Dosya", m_Dosya, m_def_Dosya)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get OtomatikCalistir() As Boolean
    OtomatikCalistir = m_OtomatikCalistir
End Property

Public Property Let OtomatikCalistir(ByVal New_OtomatikCalistir As Boolean)
    m_OtomatikCalistir = New_OtomatikCalistir
    PropertyChanged "OtomatikCalistir"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Dosya() As String
    Dosya = m_Dosya
End Property

Public Property Let Dosya(ByVal New_Dosya As String)
    m_Dosya = New_Dosya
    PropertyChanged "Dosya"
    tolgahanDosyaBilgiAl
End Property

Private Sub tolgahanDosyaBilgiAl()
    DestroyIcon hIcon
    Dim fi As SHFILEINFO
    SHGetFileInfo m_Dosya, 0, fi, Len(fi), SHGFI_ICON
    hIcon = fi.hIcon
    isim = Mid(m_Dosya, InStrRev(m_Dosya, "\") + 1)
    tolgahanDugmeCiz
End Sub
Private Sub UserControl_Click()
    If m_OtomatikCalistir Then calistir
    RaiseEvent Tikla
End Sub

