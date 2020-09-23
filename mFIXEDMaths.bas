Attribute VB_Name = "mFIXEDMaths"
Option Explicit

Public Function FixedFromDouble(ByVal d As Double) As FIXED

    ' ====================================================================================
    ' This subrountine created by...
    '
    ' Tim Arheit
    ' WordArt Control (similar project, different focus.)
    ' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37388&lngWId=1
    ' ====================================================================================
    
   Dim f As FIXED
   Dim i As Long
   
   ' Calculate the Value portion
   ' Note: -1.2 must be rounded to -2,  The Value portion can be
   ' positive or negative, but the Fract portion can only be
   ' positive.  Hence -1.2 is stored as -2 + 0.8
   i = Int(d)
   If i < 0 Then
      f.Value = &H8000 Or CInt(i And &H7FFF)
   Else
      f.Value = CInt(i And &H7FFF)
   End If
      
   i = (CLng(d * 65536#) And &HFFFF&)
   If (i And &H8000&) = &H8000& Then
      f.fract = &H8000 Or CInt(i And &H7FFF&)
   Else
      f.fract = CInt(i And &H7FFF&)
   End If
   
   FixedFromDouble = f
   
End Function

Public Function DoubleFromFixed(f As FIXED) As Double
    
    ' ====================================================================================
    ' This subrountine created by...
    '
    ' Tim Arheit
    ' WordArt Control (similar project, different focus.)
    ' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37388&lngWId=1
    ' ====================================================================================
   
   Dim d As Double
   
   d = CDbl(f.Value)
   
   If f.fract < 0 Then
      d = d + (32768 + (f.fract And &H7FFF)) / 65536#
   Else
      d = d + CDbl(f.fract) / 65536#
   End If
   
   DoubleFromFixed = d
   
End Function

