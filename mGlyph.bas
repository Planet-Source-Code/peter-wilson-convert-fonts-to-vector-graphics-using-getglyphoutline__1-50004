Attribute VB_Name = "mGlyph"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type FIXED
    fract As Integer
    Value As Integer
End Type

Public Type POINTFX
    X As FIXED
    Y As FIXED
End Type

Public Type TTPOLYCURVE
    wType As Integer
    cpfx As Integer
    apfx As POINTFX
End Type

Public Type TTPOLYGONHEADER
    cb As Long
    dwType As Long
    pfxStart As POINTFX
End Type


' The 'GetGlyphOutline' function retrieves the curve data points in
' the rasterizer's native format and uses the font's design units.
Public Const GGO_NATIVE = 2&

' The GLYPHMETRICS structure contains information about the
' placement and orientation of a glyph in a character cell.
Public Type GLYPHMETRICS
    gmBlackBoxX As Long
    gmBlackBoxY As Long
    gmptGlyphOrigin As POINTAPI
    gmCellIncX As Integer
    gmCellIncY As Integer
End Type

' The MAT2 structure contains the values for a transformation
' matrix used by the GetGlyphOutline function.
Public Type MAT2
    eM11 As FIXED
    eM12 As FIXED
    eM21 As FIXED
    eM22 As FIXED
End Type

' The GetGlyphOutline function retrieves the outline or bitmap for a character in the TrueType font that is selected into the specified device context.
Public Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hDC&, ByVal uChar&, ByVal fuFormat&, lpgm As GLYPHMETRICS, ByVal cbBuffer&, lpBuffer As Any, lpmat2 As MAT2) As Long

Public Const GDI_ERROR As Long = &HFFFF

Public Const TT_PRIM_LINE = 1       ' Curve is a polyline.
Public Const TT_PRIM_QSPLINE = 2    ' Curve is a quadratic Bézier spline.
Public Const TT_PRIM_CSPLINE = 3    ' Curve is a cubic Bézier spline.
  

' Font Points.
' =========
Public m_intPointCount As Integer
Public Type mdrFontPoint
    X As Double
    Y As Double
    Style As Integer
End Type
Public m_objFontPoints() As mdrFontPoint

' Polygons.
' =========
Public Type mdrPolygon
    Vertex() As mdrFontPoint
End Type
Public m_objPolygons() As mdrPolygon


Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

