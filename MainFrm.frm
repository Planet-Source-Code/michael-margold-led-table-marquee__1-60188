VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "LED Table Marquee"
   ClientHeight    =   330
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   540
      Top             =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Const BrightYellow = &HFFFF&
Const DarkYellow = &H6060&
Dim mDC As Long
Dim mBitmap As Long
Dim nDC As Long
Dim nBitmap As Long
Dim TotalStringLength As Long
Const CharWidth = 6
Const CharHeight = 8
Private Sub Form_Load()
  'Load font to buffer
  LoadFontContentToBuffer
  'Set the graphic mode to persistent
  Me.AutoRedraw = True
  'API uses pixels
  Me.ScaleMode = vbPixels
  'Create a device context, compatible with the screen
  mDC = CreateCompatibleDC(GetDC(0))
  'Create a bitmap, compatible with the screen
  mBitmap = CreateCompatibleBitmap(GetDC(0), FontContentSize * 3 - 1, CharHeight * 3 - 1)
  'Select the bitmap into the device context
  SelectObject mDC, mBitmap
  'Create a device context, compatible with the screen
  nDC = CreateCompatibleDC(GetDC(0))
  'Create a bitmap, compatible with the screen
  nBitmap = CreateCompatibleBitmap(GetDC(0), FontContentSize * 3 - 1, CharHeight * 3 - 1)
  'Select the bitmap into the device context
  SelectObject nDC, nBitmap
  Dim i As Long, j As Long, k As Long, l As Long
  For i = 0 To (FontContentSize - 1)
    For j = 0 To (CharHeight - 1)
      For k = 0 To 1
        For l = 0 To 1
          If FontContent(i) \ (2 ^ (7 - j)) Mod 2 = 1 Then
            SetPixel mDC, (i * 3) + k, (j * 3) + l, BrightYellow
          Else
            SetPixel mDC, (i * 3) + k, (j * 3) + l, DarkYellow
          End If
          SetPixel mDC, (i * 3) + k, (j * 3) + 2, &H0
          SetPixel mDC, (i * 3) + 2, (j * 3) + l, &H0
          SetPixel mDC, (i * 3) + 2, (j * 3) + 2, &H0
        Next l
      Next k
    Next j
  Next i
End Sub
Public Sub LedTableShow(Str As String)
  Dim i As Long
  TotalStringLength = Len(Str) * CharWidth * 3
  For i = 0 To Len(Str) - 1
    Select Case Asc(Mid(Str, i + 1, 1))
      Case 32 To 126
        BitBlt nDC, i * (CharWidth * 3), 0, (CharWidth * 3), CharHeight * 3 - 1, mDC, (Asc(Mid(Str, i + 1, 1)) - 32) * (CharWidth * 3), 0, vbSrcCopy
      Case Else
        BitBlt nDC, i * (CharWidth * 3), 0, (CharWidth * 3), CharHeight * 3 - 1, mDC, 0, 0, vbSrcCopy
    End Select
  Next i
End Sub
Private Sub Form_Resize()
  If Me.Height <> 690 Then Me.Height = 690
End Sub
Private Sub Timer1_Timer()
  Static i As Long
  i = i - 3
  If TotalStringLength > 0 Then
    If i < TotalStringLength Then i = i + TotalStringLength
    i = i Mod TotalStringLength
  End If
  LedTableShow "*** Date: " & Date & " Time: " & Time & " Created by Michael Margold (www.soft-collection.com) ***"
  Cls
  BitBlt Me.hdc, i, 0, TotalStringLength - i, CharHeight * 3 - 1, nDC, 0, 0, vbSrcCopy
  BitBlt Me.hdc, 0, 0, i, CharHeight * 3 - 1, nDC, TotalStringLength - i, 0, vbSrcCopy
  Me.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
  DeleteDC mDC
  DeleteObject mBitmap
  DeleteDC nDC
  DeleteObject nBitmap
End Sub

