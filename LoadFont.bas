Attribute VB_Name = "LoadFont"
Option Explicit

Public FontContent() As Byte
Public FontContentSize As Long
Public Sub LoadFontContentToBuffer()
  FontContent = LoadResData(101, "BIN")
  FontContentSize = UBound(FontContent) + 1 '0..x = x+1
End Sub
