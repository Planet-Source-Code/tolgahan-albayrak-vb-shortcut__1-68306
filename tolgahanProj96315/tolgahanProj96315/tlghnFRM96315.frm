VERSION 5.00
Begin VB.Form tlghnFRM96315 
   AutoRedraw      =   -1  'True
   Caption         =   "Ýstediðiniz dosylarý seçin, sürükleyin ve form alanýna býrakýn..."
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   518
   StartUpPosition =   3  'Windows Default
   Begin Project1.tlghnUC96315 tlghnUC96315 
      Height          =   1350
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   2381
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      BackColor       =   16777215
   End
End
Attribute VB_Name = "tlghnFRM96315"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim d, son As Integer
For Each d In Data.Files
    son = tlghnUC96315.UBound + 1
    Load tlghnUC96315(son)
    tlghnUC96315(son).Dosya = d
    tlghnUC96315(son).Visible = True
Next
tolgahanSimgeleriYenile
End Sub
Sub tolgahanSimgeleriYenile()
Dim i As Integer
Dim x As Integer, tmpX As Integer, y As Integer
x = Int(ScaleWidth / (tlghnUC96315(0).Width + 5))
For i = tlghnUC96315.LBound To tlghnUC96315.UBound
    If Not tlghnUC96315(i) Is Nothing Then
        If tlghnUC96315(i).Visible Then
           If tmpX = x Then tmpX = 0: y = y + 1
           tlghnUC96315(i).Left = tmpX * (tlghnUC96315(0).Width + 5)
           tlghnUC96315(i).Top = y * (tlghnUC96315(0).Height + 5)
           tmpX = tmpX + 1
        End If
    End If
Next
End Sub

Private Sub Form_Resize()
tolgahanSimgeleriYenile
End Sub
