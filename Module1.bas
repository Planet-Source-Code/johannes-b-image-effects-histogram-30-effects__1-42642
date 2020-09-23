Attribute VB_Name = "Module1"
Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

