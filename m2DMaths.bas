Attribute VB_Name = "m2DMaths"
Option Explicit

Private Const g_sngPIDivideBy180 As Single = 0.0174533!
Public Function ConvertDeg2Rad(Degress As Single) As Single

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (g_sngPIDivideBy180)
    
End Function

Public Function MatrixRotationZ(Radians As Single) As MAT2

    ' In this VB application:
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points into the monitor.
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Round(Cos(Radians), 2)
    sngSine = Round(Sin(Radians), 2)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationZ = MatrixIdentity()
    '
    ' Z-Axis rotation.
    ' A positive rotation of 90° transforms the X axis into the Y axis
    ' =================================================================
    With MatrixRotationZ
        ' Actual Rotation values.
        .eM11 = FixedFromDouble(sngCosine)
        .eM21 = FixedFromDouble(sngSine)
        .eM12 = FixedFromDouble(-sngSine)
        .eM22 = FixedFromDouble(sngCosine)
        
        ' Increase Resolution by multiplying the matrix with 256.
        .eM11 = FixedFromDouble(DoubleFromFixed(.eM11) * 256)
        .eM21 = FixedFromDouble(DoubleFromFixed(.eM21) * 256)
        .eM12 = FixedFromDouble(DoubleFromFixed(.eM12) * 256)
        .eM22 = FixedFromDouble(DoubleFromFixed(.eM22) * 256)
        
    End With

End Function
Public Function MatrixIdentity() As MAT2

    ' The identity matrix is used as the starting point for matrices
    ' that will modify vertex values to create rotations, translations,
    ' and any other transformations that can be represented by a 2x2 matrix.
    '
    ' Notice that...
    '   * the 1's go diagonally down?
    '   * rc stands for Row Column. Therefore, rc12 means Row1, Column 2.
    
    With MatrixIdentity
        ' Value Part
        .eM11.Value = 1: .eM12.Value = 0
        .eM21.Value = 0: .eM22.Value = 1
        
        ' Fraction Part
        .eM11.fract = 0: .eM12.fract = 0
        .eM21.fract = 0: .eM22.fract = 0
    End With
    
End Function


