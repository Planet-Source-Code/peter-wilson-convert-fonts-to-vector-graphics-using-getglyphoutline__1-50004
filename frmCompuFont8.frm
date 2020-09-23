VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCompuFont8 
   AutoRedraw      =   -1  'True
   Caption         =   "CompuFont8"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "frmCompuFont8.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnHelp 
      Cancel          =   -1  'True
      Caption         =   "&Help"
      Height          =   405
      Left            =   60
      TabIndex        =   13
      Top             =   6210
      Width           =   1245
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   5910
      TabIndex        =   14
      Top             =   6210
      Width           =   1245
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Output"
      Height          =   2595
      Left            =   60
      TabIndex        =   10
      Top             =   3540
      Width           =   7095
      Begin VB.PictureBox pictRGBTicks 
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1260
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   22
         Top             =   2130
         Width           =   345
      End
      Begin VB.PictureBox pictRGBStart 
         BackColor       =   &H0000FFFF&
         Height          =   345
         Left            =   900
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   21
         Top             =   2130
         Width           =   345
      End
      Begin VB.PictureBox pictRGBOutline 
         BackColor       =   &H000000FF&
         Height          =   345
         Left            =   510
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   20
         Top             =   2130
         Width           =   345
      End
      Begin VB.PictureBox pictRGBBackground 
         BackColor       =   &H00000000&
         Height          =   345
         Left            =   120
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   19
         Top             =   2130
         Width           =   345
      End
      Begin RichTextLib.RichTextBox rtfOutput 
         Height          =   2205
         Left            =   1680
         TabIndex        =   12
         Top             =   270
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   3889
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmCompuFont8.frx":044A
      End
      Begin VB.PictureBox pictOutput 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   120
         ScaleHeight     =   1425
         ScaleWidth      =   1425
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Width           =   1485
      End
   End
   Begin VB.Frame frameUserInput 
      Caption         =   "User Input"
      Height          =   2205
      Left            =   30
      TabIndex        =   0
      Top             =   1260
      Width           =   7095
      Begin MSComctlLib.Slider Slider1 
         Height          =   225
         Left            =   30
         TabIndex        =   23
         ToolTipText     =   "Rotate around Z-axis."
         Top             =   1890
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   45
         Min             =   -180
         Max             =   180
         TickFrequency   =   45
      End
      Begin VB.CheckBox chkDrawSlowly 
         Caption         =   "Draw Slowly"
         Height          =   225
         Left            =   5730
         TabIndex        =   9
         Top             =   1590
         Width           =   1275
      End
      Begin VB.CheckBox chkShowVertices 
         Caption         =   "Show Vertices"
         Height          =   225
         Left            =   3045
         TabIndex        =   8
         Top             =   1590
         Width           =   1425
      End
      Begin VB.CheckBox chkShowStart 
         Caption         =   "Show Start Vertex"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   1590
         Width           =   1665
      End
      Begin VB.TextBox txtSplineResolution 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3915
      End
      Begin VB.TextBox txtFont 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   705
         Width           =   5295
      End
      Begin MSComctlLib.Slider SliderSplineResolution 
         Height          =   285
         Left            =   5580
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   1
         Max             =   6
         SelStart        =   2
         Value           =   2
      End
      Begin VB.TextBox txtUserInput 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   6825
      End
      Begin VB.CommandButton btnChooseFont 
         Caption         =   "Choose &Font..."
         Height          =   315
         Left            =   5610
         TabIndex        =   3
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label lblSplineResolution 
         AutoSize        =   -1  'True
         Caption         =   "Spline Resolution"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1125
         Width           =   1230
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7650
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pictTop 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   7125
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   7185
      Begin VB.Image Image1 
         Height          =   240
         Left            =   5400
         Picture         =   "frmCompuFont8.frx":04E2
         ToolTipText     =   "Supports the Asian Language"
         Top             =   870
         Width           =   1680
      End
      Begin VB.Image imgLogo 
         Height          =   855
         Left            =   120
         Picture         =   "frmCompuFont8.frx":09DF
         Top             =   120
         Width           =   795
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (c) 2003 Peter Wilson, http://dev.midar.com/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   210
         Index           =   2
         Left            =   1050
         TabIndex        =   18
         Top             =   780
         Width           =   3885
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Convert any True-Type Font into Polygons."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   285
         Index           =   1
         Left            =   1020
         TabIndex        =   17
         Top             =   480
         Width           =   4920
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CompuFont8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   1020
         TabIndex        =   16
         Top             =   60
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmCompuFont8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' URL's that were helpful in this program's creation.
' (I found this a very confusing topic! Not all of these URL's were helpful.
' ==========================================================================
'   * http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnaskdr/html/drgui55.asp
'   * Microsoft Knowledge Base Article - 243285. HOWTO: Draw TrueType Glyph Outlines
'   * http://groups.google.com.au/groups?hl=en&lr=&ie=UTF-8&oe=UTF-8&threadm=3AF43F13.EDF682C9%40gmx.de-REMOVE&rnum=3&prev=/groups%3Fhl%3Den%26ie%3DUTF-8%26oe%3DUTF-8%26q%3DTTPOLYGONHEADER%2B%2BVB%26sa%3DN%26tab%3Dwg%26meta%3D
'   * C++ Example Code: http://my.execpc.com/~dg/tutorial/Glyph/Glyph.html
'   * http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37388&lngWId=1
'   * http://groups.google.com.au/groups?hl=en&lr=&ie=UTF-8&oe=UTF-8&frame=right&th=65d462ec549a0592&seekm=1992Apr22.043348.3671%40microsoft.com#link1
'
'   Most helpful URL's
'   * ms-help://MS.VSCC.2003/MS.MSDNQTR.2003FEB.1033/gdi/fontext_52at.htm
'   * http://support.microsoft.com/default.aspx?scid=http://support.microsoft.com:80/support/kb/articles/q87/1/15.asp&NoWebContent=1
'   * http://dev.midar.com/ (my personal web site)

Private sngResolution As Single

Private lngOutlineColour As Long

Private Sub Spline(PointA As mdrFontPoint, PointB As mdrFontPoint, PointC As mdrFontPoint)
    
    Dim sngT As Single
    Dim sngX As Single
    Dim sngY As Single
    
    For sngT = 0 To 1 Step sngResolution
        sngX = (PointA.X - 2 * PointB.X + PointC.X) * sngT ^ 2 + (2 * PointB.X - 2 * PointA.X) * sngT + PointA.X
        sngY = (PointA.Y - 2 * PointB.Y + PointC.Y) * sngT ^ 2 + (2 * PointB.Y - 2 * PointA.Y) * sngT + PointA.Y
        
        m_intPointCount = m_intPointCount + 1
        ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
                       m_objFontPoints(m_intPointCount).Style = TT_PRIM_QSPLINE
                       m_objFontPoints(m_intPointCount).X = sngX
                       m_objFontPoints(m_intPointCount).Y = sngY
    Next sngT
        
End Sub

Private Sub GetGlyphs(ByVal hDC&, WhichByte As Integer, sngScale As Single)
    
    ' =================================================================
    ' Thanks to "James Ho" for suggesting improvements to this routine,
    ' so that it now supports chinese characters as input.
    ' =================================================================
    
    Dim lngTotalNativeBuffer As Long
    Dim lpgm As GLYPHMETRICS
    Dim lpmat2 As MAT2
    Dim abytBuffer() As Byte
    
    
    Dim intN As Integer
    Dim objPolyHeader() As TTPOLYGONHEADER
    Dim objPolyCurve() As TTPOLYCURVE
    Dim objPointFX() As POINTFX
    Dim lngTotalPolygonHeader As Long
    Dim lngIndex As Long
    Dim lngIndexPolygonHeader As Long
    Dim lngPolygonCount As Long
    
    lpmat2 = MatrixRotationZ(ConvertDeg2Rad(Me.Slider1.Value))

    ' Get the required buffer size  (This works)
    lngTotalNativeBuffer = GetGlyphOutline(hDC, WhichByte, GGO_NATIVE, lpgm, 0, ByVal 0&, lpmat2)
    If lngTotalNativeBuffer <> GDI_ERROR Then

        ' Set the buffer size
        ReDim abytBuffer(lngTotalNativeBuffer - 1) As Byte
        
        'Then retrieve the information
        If GetGlyphOutline(hDC, WhichByte, GGO_NATIVE, lpgm, lngTotalNativeBuffer, abytBuffer(0), lpmat2) <> GDI_ERROR Then
            
''            Debug.Print lpgm.gmptGlyphOrigin.X
''            Debug.Print lpgm.gmBlackBoxX
''            Debug.Print lpgm.gmBlackBoxY
''            Debug.Print lpgm.gmptGlyphOrigin.Y
''
''            Debug.Print lpgm.gmCellIncX
''            Debug.Print lpgm.gmCellIncY

            Me.pictOutput.ScaleLeft = lpgm.gmptGlyphOrigin.X
            Me.pictOutput.ScaleWidth = lpgm.gmBlackBoxX
            Me.pictOutput.ScaleHeight = lpgm.gmBlackBoxY
            Me.pictOutput.ScaleTop = -lpgm.gmptGlyphOrigin.Y
            
            frmPreview.pictOutput.ScaleLeft = lpgm.gmptGlyphOrigin.X
            frmPreview.pictOutput.ScaleWidth = lpgm.gmBlackBoxX
            frmPreview.pictOutput.ScaleHeight = lpgm.gmBlackBoxY
            frmPreview.pictOutput.ScaleTop = -lpgm.gmptGlyphOrigin.Y
            
            
            ReDim objPolyHeader(0) As TTPOLYGONHEADER
            ReDim objPolyCurve(0) As TTPOLYCURVE

            lngIndex = 0
            m_intPointCount = 0
            
            Do
                ' ==================================================================
                ' Copy a PolygonHeader into memory (and increment the buffer index).
                ' ==================================================================
                CopyMemory objPolyHeader(0), abytBuffer(lngIndex), 16
                lngTotalPolygonHeader = objPolyHeader(0).cb
                lngIndex = lngIndex + 16
                lngIndexPolygonHeader = 16
                
                ' ===================================================================================
                ' A PolygonHeader has a start point, that can either be the start of a straight line,
                ' or the start of a quadratic Bézier spline. i.e. Point A, in an [A,B,C] curve.
                ' ===================================================================================
                m_intPointCount = m_intPointCount + 1
                ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
                               m_objFontPoints(m_intPointCount).Style = 0 ' ie. Starting point.
                               m_objFontPoints(m_intPointCount).X = DoubleFromFixed(objPolyHeader(0).pfxStart.X)
                               m_objFontPoints(m_intPointCount).Y = DoubleFromFixed(objPolyHeader(0).pfxStart.Y)
                               
                Do
                    ' ======================================================
                    ' At least one PolyCurve always follows a PolygonHeader.
                    ' PolyCurve always has a least one starting PointFX.
                    ' ======================================================
                    CopyMemory objPolyCurve(0), abytBuffer(lngIndex), 12
                    lngIndex = lngIndex + 12
                    lngIndexPolygonHeader = lngIndexPolygonHeader + 12
                    
                    ' =======================================
                    ' Load additional PointFX values (if any).
                    ' =======================================
                    If objPolyCurve(0).cpfx > 1 Then
                        ReDim objPointFX((objPolyCurve(0).cpfx - 2)) As POINTFX
                        CopyMemory objPointFX(0), abytBuffer(lngIndex), (8 * (objPolyCurve(0).cpfx - 1))
                        lngIndex = lngIndex + (8 * (objPolyCurve(0).cpfx - 1))
                        lngIndexPolygonHeader = lngIndexPolygonHeader + (8 * (objPolyCurve(0).cpfx - 1))
                    End If
                
                    ' Part A) Create the initial polycurve point...
                    m_intPointCount = m_intPointCount + 1
                    ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
                                   m_objFontPoints(m_intPointCount).Style = TT_PRIM_LINE
                                   m_objFontPoints(m_intPointCount).X = DoubleFromFixed(objPolyCurve(0).apfx.X)
                                   m_objFontPoints(m_intPointCount).Y = DoubleFromFixed(objPolyCurve(0).apfx.Y)
                
                    ' Post-Process points depending on whether they are straight lines, or curves.
                    ' ===========================================================================
                    If objPolyCurve(0).wType = TT_PRIM_LINE Then
    
                        ' ============================
                        ' PointFX(0..n) is a polyline.
                        ' ============================
                        
                        ' Part B) ...Create subsequent points.
                        If objPolyCurve(0).cpfx > 1 Then
                            For intN = LBound(objPointFX) To UBound(objPointFX)
    
                                m_intPointCount = m_intPointCount + 1
                                ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
                                               m_objFontPoints(m_intPointCount).Style = TT_PRIM_LINE
                                               m_objFontPoints(m_intPointCount).X = DoubleFromFixed(objPointFX(intN).X)
                                               m_objFontPoints(m_intPointCount).Y = DoubleFromFixed(objPointFX(intN).Y)
                            Next intN
                        End If
    
                    ElseIf objPolyCurve(0).wType = TT_PRIM_QSPLINE Then
                        ' ======================================================================
                        ' PointFX(0..n) is a quadratic Bézier spline.
                        ' Load the Spline's Control points first, then create the Spline itself.
                        ' ======================================================================
                        Dim lngControlPoint  As Long
                        Dim objSplinePoints() As mdrFontPoint
                        ReDim objSplinePoints(1) As mdrFontPoint
                        
                        ' The last defined point is Point A (from a previous object/header).
                        ' ==================================================================
                        lngControlPoint = 0
                        objSplinePoints(0).Style = -1 ' ie. Control Point for spline.
                        objSplinePoints(0).X = m_objFontPoints(m_intPointCount - 1).X
                        objSplinePoints(0).Y = m_objFontPoints(m_intPointCount - 1).Y
                        
                        ' This is the first control point B.
                        lngControlPoint = lngControlPoint + 1
                        objSplinePoints(1).Style = -1 ' ie. Control Point for spline.
                        objSplinePoints(1).X = m_objFontPoints(m_intPointCount).X
                        objSplinePoints(1).Y = m_objFontPoints(m_intPointCount).Y
                        
                        If objPolyCurve(0).cpfx > 1 Then ' Splines always have a number greater than 1.
                            For intN = LBound(objPointFX) To UBound(objPointFX)
                            
                                lngControlPoint = lngControlPoint + 1
                                ReDim Preserve objSplinePoints(lngControlPoint) As mdrFontPoint
                                               objSplinePoints(lngControlPoint).Style = -1 ' ie. Control Point for spline.
                                               objSplinePoints(lngControlPoint).X = DoubleFromFixed(objPointFX(intN).X)
                                               objSplinePoints(lngControlPoint).Y = DoubleFromFixed(objPointFX(intN).Y)
                            Next intN
                        End If
                        
                        If sngResolution > 0 Then ' Do Curved surfaces...
                        
                            ' ========================================================================================
                            ' At this point, we have an array of control points that help make up a series of splines.
                            ' ie. objSplinePoints(0 to lngControlPoint)
                            ' Note: They may be in the form A,B,B,B,C.
                            '       In which case, new C's need to be found inbetween the B's (except the last one).
                            ' ========================================================================================
                            Dim intB As Integer
                            Dim PointA As mdrFontPoint
                            Dim PointB As mdrFontPoint
                            Dim PointC As mdrFontPoint
                            
                            intB = 1
                            PointC = objSplinePoints(intB - 1)
                            
                            lngControlPoint = lngControlPoint + 1
                            ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
                                           m_objFontPoints(m_intPointCount).Style = -1 ' ie. Control Point for spline.
                                           m_objFontPoints(m_intPointCount).X = PointC.X
                                           m_objFontPoints(m_intPointCount).Y = PointC.Y
                            
                            Do
                                PointA = PointC
                                PointB = objSplinePoints(intB)
                                If intB < (lngControlPoint - 2) Then
                                    ' Find a new midpoint
                                    PointC.X = (objSplinePoints(intB).X + objSplinePoints(intB + 1).X) / 2
                                    PointC.Y = (objSplinePoints(intB).Y + objSplinePoints(intB + 1).Y) / 2
                                Else
                                    PointC = objSplinePoints(intB + 1)
                                End If
                                
                                Call Spline(PointA, PointB, PointC)
                                intB = intB + 1
                                
                            Loop Until intB = (lngControlPoint - 1)
                        End If ' Ignore curved surfaces if sngResolution = 0.
                        
                    Else
                        ' Unsupport style of font.
                        Err.Raise vbObjectError + 1001, "GetGlyphs", "Unsupported Font Style. Please contact technical support (http://dev.midar.com) for additional information on this error."
                    End If
                
                Loop Until lngIndexPolygonHeader = lngTotalPolygonHeader
            Loop Until lngIndex = lngTotalNativeBuffer


            ' =====================================================================
            ' Clean up, Remove Duplicates, and separate into Polygons and Vertices.
            ' =====================================================================
            Dim intVertexCount As Integer
            Dim intPolyCount As Integer
            intPolyCount = -1
            For intN = LBound(m_objFontPoints) + 1 To UBound(m_objFontPoints)
                If m_objFontPoints(intN).Style = 0 Then
                    intVertexCount = -1
                    intPolyCount = intPolyCount + 1
                    ReDim Preserve m_objPolygons(intPolyCount)
                End If
                
                If (m_objFontPoints(intN - 1).X = m_objFontPoints(intN).X) And _
                   (m_objFontPoints(intN - 1).Y = m_objFontPoints(intN).Y) Then
                    ' Do nothing because this is a duplicate item.
                Else
                    ' Add new vertex.
                    intVertexCount = intVertexCount + 1
                    ReDim Preserve m_objPolygons(intPolyCount).Vertex(intVertexCount)
                                   m_objPolygons(intPolyCount).Vertex(intVertexCount).Style = m_objFontPoints(intN).Style
                                   m_objPolygons(intPolyCount).Vertex(intVertexCount).X = m_objFontPoints(intN).X
                                   m_objPolygons(intPolyCount).Vertex(intVertexCount).Y = m_objFontPoints(intN).Y
                End If
                
            Next intN


        Else
            Err.Raise Err.LastDllError
        End If
    Else
        Err.Raise Err.LastDllError
    End If
    
End Sub



Private Sub btnChooseFont_Click()

    On Error GoTo errTrap
    
    Me.CommonDialog1.Flags = cdlCFBoth Or cdlCFTTOnly
    Me.CommonDialog1.ShowFont
    
    If Trim(Me.CommonDialog1.FontName) = "" Then Exit Sub
    
    Me.txtFont.Text = Me.CommonDialog1.FontName
    Me.pictOutput.FontName = Me.CommonDialog1.FontName
    Me.pictOutput.FontBold = Me.CommonDialog1.FontBold
    Me.pictOutput.FontItalic = Me.CommonDialog1.FontItalic
    Me.pictOutput.FontSize = Me.CommonDialog1.FontSize
    
    Call txtUserInput_Change
    
    Exit Sub
errTrap:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub DrawGlyph(Canvas As PictureBox, WriteFile As Boolean, Optional DrawStart As Boolean = False, Optional DrawTicks As Boolean = False, Optional DrawSlowly As Boolean = False)

    On Error GoTo errTrap
    
    Dim intPoly As Integer
    Dim intVertex As Integer
    Dim sngX As Single
    Dim sngY As Single
    
    Canvas.Cls
    Canvas.BackColor = Me.pictRGBBackground.BackColor
        
    Me.rtfOutput.Text = ""
    
    ' Create a Temporary File in the user's temp directory.
    ' =====================================================
    Dim strTempFileName As String
    If WriteFile = True Then
        strTempFileName = GetWindowsTempFile
        Open strTempFileName For Output As #1
    End If
    
    For intPoly = LBound(m_objPolygons) To UBound(m_objPolygons)
        If WriteFile = True Then
            Print #1, "Polygon " & intPoly & "/" & UBound(m_objPolygons) & vbCrLf & "{"
        End If
        
        For intVertex = LBound(m_objPolygons(intPoly).Vertex) To UBound(m_objPolygons(intPoly).Vertex)
            
            ' Default drawing styles (can be overridden).
            ' ===========================================
            Canvas.DrawWidth = 1
            Canvas.ForeColor = Me.pictRGBOutline.BackColor
            
            sngX = m_objPolygons(intPoly).Vertex(intVertex).X
            sngY = -m_objPolygons(intPoly).Vertex(intVertex).Y
            
            If intVertex = LBound(m_objPolygons(intPoly).Vertex) Then
                If (DrawStart = True) Then
                    Canvas.DrawWidth = 3
                    Canvas.ForeColor = Me.pictRGBStart.BackColor
                End If
                Canvas.PSet (sngX, sngY)
            Else
                Canvas.Line -(sngX, sngY)
            End If
            
            ' ===========
            ' Draw Ticks.
            ' ===========
            If (DrawTicks = True) Then
                Canvas.DrawWidth = 2
                Canvas.ForeColor = Me.pictRGBTicks.BackColor
                Canvas.PSet (sngX, sngY)
            End If
            
            If DrawSlowly = True Then Canvas.Refresh

            If WriteFile = True Then
                Print #1, "    Vertex " & intVertex & "/" & UBound(m_objPolygons(intPoly).Vertex) + 1 & " {x:" & sngX & ", y:" & sngY & "}"
            End If

        Next intVertex
        
        ' Close the polygon.
        Canvas.DrawWidth = 1
        Canvas.ForeColor = Me.pictRGBOutline.BackColor
        sngX = m_objPolygons(intPoly).Vertex(0).X
        sngY = -m_objPolygons(intPoly).Vertex(0).Y
        Canvas.Line -(sngX, sngY)
        If WriteFile = True Then
            Print #1, "    Vertex " & intVertex & "/" & UBound(m_objPolygons(intPoly).Vertex) + 1 & " {x:" & sngX & ", y:" & sngY & "}" & vbCrLf & "}" & vbCrLf
        End If

    Next intPoly
    
    If WriteFile = True Then
        Close #1
        Me.rtfOutput.LoadFile (strTempFileName)
        ' Remove the temporary file (that was created at the top of this subroutine)
        Kill strTempFileName
    End If
    
    
    Exit Sub
errTrap:
    MsgBox Err.Number & " - " & Err.Description, vbCritical
    
End Sub

   
Private Sub btnHelp_Click()
    
    Call ShowHTMLHelp(Me.hWnd)
    
End Sub

Private Sub chkDrawSlowly_Click()
    Call txtUserInput_Change
End Sub

Private Sub chkShowStart_Click()
    Call txtUserInput_Change
End Sub

Private Sub chkShowVertices_Click()
    Call txtUserInput_Change
End Sub

Private Sub Form_Load()

    lngOutlineColour = RGB(255, 0, 0)
    
    Call SliderSplineResolution_Scroll
    
    Me.txtFont.Text = Me.pictOutput.FontName
    Me.pictOutput.FontName = Me.pictOutput.FontName
    Me.pictOutput.FontBold = Me.pictOutput.FontBold
    Me.pictOutput.FontItalic = Me.pictOutput.FontItalic
    Me.pictOutput.FontSize = Me.pictOutput.FontSize
    
    frmPreview.Show vbModeless, Me
    
End Sub

Private Sub pictRGBBackground_Click()
    
    Me.CommonDialog1.ShowColor
    Me.pictRGBBackground.BackColor = Me.CommonDialog1.Color
    Call txtUserInput_Change
    
End Sub

Private Sub pictRGBOutline_Click()

    Me.CommonDialog1.ShowColor
    Me.pictRGBOutline.BackColor = Me.CommonDialog1.Color
    Call txtUserInput_Change
    
End Sub


Private Sub pictRGBStart_Click()

    Me.CommonDialog1.ShowColor
    Me.pictRGBStart.BackColor = Me.CommonDialog1.Color
    Call txtUserInput_Change
    
End Sub

Private Sub pictRGBTicks_Click()

    Me.CommonDialog1.ShowColor
    Me.pictRGBTicks.BackColor = Me.CommonDialog1.Color
    Call txtUserInput_Change
    
End Sub

Private Sub Slider1_Scroll()
    Call txtUserInput_Change
End Sub


Private Sub SliderSplineResolution_Scroll()

    Select Case Me.SliderSplineResolution.Value
        Case 0
            Me.txtSplineResolution.Text = "0.0  - Skip over curved surfaces completely."
            sngResolution = 0
        
        Case 1
            Me.txtSplineResolution.Text = "1.0  - Draw start & end points of curved surfaces only."
            sngResolution = 1
            
        Case 2
            Me.txtSplineResolution.Text = "0.5  - Low Resolution"
            sngResolution = 0.5
            
        Case 3
            Me.txtSplineResolution.Text = "0.25 - Normal. A good balance of speed and quality."
            sngResolution = 0.25
            
        Case 4
            Me.txtSplineResolution.Text = "0.1  - High Resolution - Suitable for CNC engraving machine output."
            sngResolution = 0.1
            
        Case 5
            Me.txtSplineResolution.Text = "0.05 - Very High Resolution."
            sngResolution = 0.05
            
        Case 6
            Me.chkDrawSlowly.Value = vbChecked
            Me.txtSplineResolution.Text = "0.01 - Massive Resolution."
            sngResolution = 0.01
            
    End Select
    
    
    Call txtUserInput_Change
    
End Sub

Private Sub txtUserInput_Change()

    On Error GoTo errTrap
    
    Dim intASCII As Integer
    Dim blnShowStartPoint As Boolean
    Dim blnShowVertices As Boolean
    Dim blnDrawSlowly As Boolean
    
    If Trim(Me.txtUserInput.Text) = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    blnShowStartPoint = False
    blnShowVertices = False
    blnDrawSlowly = False
    
    If Me.chkShowStart.Value = 1 Then blnShowStartPoint = True
    If Me.chkShowVertices.Value = 1 Then blnShowVertices = True
    If Me.chkDrawSlowly.Value = 1 Then blnDrawSlowly = True
    
    intASCII = Asc(Right(Me.txtUserInput.Text, 1))
    
    
    Call GetGlyphs(Me.pictOutput.hDC, intASCII, 256)
    Call DrawGlyph(Me.pictOutput, False, blnShowStartPoint, blnShowVertices, blnDrawSlowly)
    Call DrawGlyph(frmPreview.pictOutput, True, blnShowStartPoint, blnShowVertices, blnDrawSlowly)
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errTrap:
    Dim strMsg As String
    
    Screen.MousePointer = vbDefault
    strMsg = "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
             "Not all keyboard items have a printable character." & vbCrLf & vbCrLf & _
             "Example:" & vbCrLf & _
             "The space bar does not draw anything."
    
    rtfOutput.Text = strMsg
    

End Sub

