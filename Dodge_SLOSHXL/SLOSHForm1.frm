'
'  IMPORTANT:
'  The user is required to put the full path and name for the
'  DLL entitled SLOSH_DLLr in the Declare Sub statement below.
'  The DLL holds the code for computing the slosh model parameters
'  The data types of the arguments must not be altered from those shown below.
'
'  It is strongly recommended that the user make a backup copy
'  of this Excel spreadsheet and the DLL immediately after
'  copying the files from the distribution media.
'  This spreadsheet and the VBA code are distributed without
'  protection or passwords.
'
'
Private Declare Sub SLOSH_Calc _
   Lib "D:\Documents\AE\Research3\Code-Implementation\Dodge_SLOSHXL\SLOSH_DLLr_04.dll" _
   (Ii1 As Long, Ii2 As Long, Io1 As Long, Ri1 As Double, Ri2 As Double, Ro1 As Double, Ro2 As Double)
Option Explicit
'
'
'  SLOSH-XL Public Version
'  Sloshing analysis in axisymmetric tanks.
'  Adapted from the WINSLOSH code written by Frank Dodge
'
'  Steve Green, Southwest Reseach Institute, July 2016
'
' Notes:
'
' Version 0.1 January, 2017
' Initial distribution for user testing.
'
' Version 0.2 February, 2017
' Included specific type declarations
' Changed the DLL to pass double precision real numbers instead of single precision
'
' Version 0.3 August, 2017
' Corrected error in the reading of the baffle specifications.
' Original has the number of submerged baffles being set to the input box text for tank radius.
' Correction is to have it set to the input box text for the number of submerged baffles.
'
' Version 0.4 February, 2018
' - Inserted instruction I.2 to replace the pathname of the DLL to the path where the
'   user stored the DLL after copyin to the user's computer.
' - Rebuilt DLL to use only Windows system libraries and DLL's
' - Changed definition of PicWidth, PicHeight in DrawBox to remove an error condition
'   that occurs in Excel 2016
' - Corrected an error in the conical tank damping calculations for radians-degrees conversions.
'
'-------------------  Declarations for Module  -------------------------
Private NSegs As Integer
Private RMax As Double
Private ZMax As Double
Private ZLiquid As Double
'--------------------------- NUMBER OF SEGMENTS -----------------------------
' These dimensions should be made large enough to hold all the tank segments:
' Example: SegType(n)   n= max allowable segments
'
'
Private SegType(20) As Integer
Private RStart(20) As Double
Private ZStart(20) As Double
Private REnd(20) As Double
Private ZEnd(20) As Double
Private RO(20) As Double
Private ZO(20) As Double
Private A1(20) As Double
Private A2(20) As Double
Private A3(20) As Double
Private A4(20) As Double
Private A5(20) As Double
Private Rradius(20) As Double
Private Zradius(20) As Double
'-----------------------------------------------------------------------------
'  Added the _1 arrays to hold the wall pressure coefficeints for printing
'  These arrays are dimensioned to provide for 11 points (10 segments) along
'  the wall where pressure data are computed.
'  Plus one point at middle of a flat bottom tank, plus one more if the flat
'  bottom intersect the centerline. For a total of 13 maximum wall locations
'  where pressure info is computed.
'
Private ZP_1(13) As Double
Private RPOut_1(13) As Double
Private POut1_1(13) As Double
Private POut2_1(13) As Double
Private RPIn_1(13) As Double
Private PIn1_1(13) As Double
Private PIn2_1(13) As Double
'
'  Free Surface Parameters
Private RBar As Double
Private IIn As Integer
Private IOut As Integer
Private RLiquid(2) As Double
Private IEps(2) As Integer
Private Eps As Double
Private NOuter As Integer
Private NInner As Integer
'
' Temporary variables needed in several spots
Private NRowsCols As Integer
Private NNSegs As Integer
'
' Damping related inputs
Private IndexType_damp As Integer
Private ZL_damp As Double
Private RTank_damp As Double
Private RIn_damp As Double
Private Angle_damp As Double
Private BW_damp As Double
Private Axspc_damp As Double
Private NumSub_damp As Integer
Private TopSub_damp As Double
Private IndexBaffle As Integer
'
' Properties
Private GLevel As Double          ' Thrust, Length/Time^2
Private Density As Double         ' Fluid Density,  Mass/Lemgth^3
Private Viscosity As Double       ' Fluid Kimematic Viscosity,  Length^2/Time
'
' UserForm Parameters
Dim BasicsChk As Integer
Dim GeomDefChk As Integer
Dim PropDefChk As Integer
Dim DampDefChk As Integer
Dim ISeg1 As Integer
Dim StartSheetName As String
'
' Tank Layout Image Paramteters
Dim x0_draw As Single
Dim z0_draw As Single
Dim xzmx_draw As Single
Dim scale_draw As Single
Dim zmx_draw As Single
Dim EnvPoly(1 To 5, 1 To 2) As Single
Dim SegPoly(1 To 20 * 100, 1 To 2) As Single
Dim CtrPoly(1 To 2, 1 To 2) As Single
Dim NppDraw As Integer
Dim Fname As String
Dim PicWidth As Single
Dim PicHeight As Single
Dim PicX0 As Single
Dim PicZ0 As Single
Private LastrowIn As Integer
'
' Value of pi
Dim PiVal As Double

Private Sub BasicsOK_Click()
'
'  Process the inputs from the Basics Page
'
Dim EnvOK As Integer
Dim NameOK As Integer
'
GeomDefChk = 0
PropDefChk = 0
DampDefChk = 0
ISeg1 = 1
NppDraw = 0
PiVal = 4 * Atn(1)
StartSheetName = ActiveSheet.Name
'
' Make sure a name is entered
If SheetName.Text = "" Then
    MsgBox "You must enter a Problem Name."
    SheetName.SetFocus
    Exit Sub
End If


If Not (ReadSheet) Then
    If SheetExists(SheetName.Text) Then
       NameOK = MsgBox("This Problem name already exsist - Overwrite?", vbYesNo + vbQuestion + vbDefaultButton2)
'
' Return to Basics Page if Not OK
       If NameOK = vbNo Then
          SLOSHForm1.MultiPage1.Value = 0
          SheetName.Text = ""
          SheetName.SetFocus
          Exit Sub
       End If
    End If
End If
'-----------------------------------------------
'  For reading an existing sheet
If ReadSheet Then
' Make sure the sheet exists
    If Not (SheetExists(SheetName.Text)) Then
       MsgBox " You must enter a valid Worksheet Name"
       ReadSheet.Value = 0
       SheetName.SetFocus
       SLOSHForm1.MultiPage1.Value = 0
       Exit Sub
    End If
'    UpdateDisplay = False
    Application.ScreenUpdating = False
    Call ReadTank(SheetName)
    BasicsChk = 1
    GeomDefChk = 1
'
    DensText.Text = CStr(Density)
    ViscText.Text = CStr(Viscosity)
    GravText.Text = CStr(GLevel)
    ZLiquidText.Text = CStr(ZLiquid)
'
    If DampAnn Then
       ZFillDampAnnText.Text = CStr(ZL_damp)
       RTankDampAnnText.Text = CStr(RTank_damp)
       RTankInAnnText.Text = CStr(RIn_damp)
    End If
    If (DampCyl) Then
       ZFillDampCylText.Text = CStr(ZL_damp)
       RTankDampCylText.Text = CStr(RTank_damp)
    End If
    If (DampSpr) Then
       ZFillDampSprText.Text = CStr(ZL_damp)
       RTankDampSprText.Text = CStr(RTank_damp)
    End If
    If (DampTor) Then
        ZFillDampTorText.Text = CStr(ZL_damp)
        RTankDampTorText.Text = CStr(RTank_damp)
    End If
    If (DampCon) Then
       ZFillDampConText.Text = CStr(ZL_damp)
       AngDampConText.Text = CStr(Angle_damp) * 180 / PiVal
    End If
    If (DampBaf) Then
       ZFillDampBafText.Text = CStr(ZL_damp)
       RTankDampBafText.Text = CStr(RTank_damp)
       BWidDampBafText.Text = CStr(BW_damp)
       ASpcDampBafText.Text = CStr(Axspc_damp)
       NSubDampBafText.Text = CStr(NumSub_damp)
       TSubDampBafText.Text = CStr(TopSub_damp)
    End If
    PicWidth = Val(RMax) * scale_draw * 1.05
    PicHeight = Val(ZMax) * scale_draw * 1.05
'    UpdateDisplay = True
    Application.ScreenUpdating = True
    MsgBox " All inputs read - You can change the Fluid  Properties, Fill Level, and Damping Model"
    SLOSHForm1.MultiPage1.Value = 2
    Exit Sub
End If
'---End of reading an existing sheet ------------------------------------
'
' Make sure the Number of Segments is entered
If NSegsText.Text = "" Then
    MsgBox "You must enter the Number of Segments."
    NSegsText.SetFocus
    Exit Sub
End If
' Make sure the Rmax is entered
If RMaxText.Text = "" Then
    MsgBox "You must enter the RMax."
    RMaxText.SetFocus
    Exit Sub
End If
' Make sure the Zmax is entered
If ZMaxText.Text = "" Then
    MsgBox "You must enter the ZMax."
    ZMaxText.SetFocus
    Exit Sub
End If
' Make sure the RStart is entered
If ZStartText.Text = "" Then
    MsgBox "You must enter the ZStart."
    ZStartText.SetFocus
    Exit Sub
End If
' Make sure the Zmax is entered
If RStartText.Text = "" Then
    MsgBox "You must enter the RStart."
    RStartText.SetFocus
    Exit Sub
End If
'
' Record that there was no comment filled in
If Comment.Text = "" Then
Comment.Text = "No Comment Recorded for " & SheetName.Text
End If
'
' Open the new worksheet as a place to land to prevent screen flashing
Application.DisplayAlerts = False
If ((SheetExists(SheetName.Text))) Then
   Sheets(SheetName.Text).Delete
End If
Sheets.Add
ActiveSheet.Select
ActiveSheet.Name = SheetName.Text
ActiveSheet.Move after:=Worksheets(Worksheets.Count)
ActiveSheet.Tab.ColorIndex = xlColorIndexNone
Range("A1").Select
Application.DisplayAlerts = True
'
'  Sketch the overall tank envelope rectangle
PicWidth = Val(RMaxText) * scale_draw * 1.05
PicHeight = Val(ZMaxText) * scale_draw * 1.05
Fname = DrawBox(Val(RMaxText), Val(ZMaxText))
Application.ScreenUpdating = True
'
' Make the layout page visible and show current sketch
SLOSHForm1.MultiPage1.Pages(4).Visible = True
SLOSHForm1.TankShape.Width = PicWidth
SLOSHForm1.TankShape.Height = PicHeight
SLOSHForm1.TankShape.Visible = True
SLOSHForm1.TankShape.Picture = LoadPicture(Fname)
SLOSHForm1.MultiPage1.Value = 4
'
' Check if the Tank extents are OK
EnvOK = MsgBox("Is the tank envelope OK?", vbYesNo + vbQuestion + vbDefaultButton2)
'
' Hide the layout page
SLOSHForm1.MultiPage1.Pages(4).Visible = False
SLOSHForm1.TankShape.Visible = False
'
' Return to Basics Page if Not OK
If EnvOK = vbNo Then
   SLOSHForm1.MultiPage1.Value = 0
   RMaxText.SetFocus
   Exit Sub
End If
'
'  Set the Basics OK Flag, transfer numericla values and go to next page
BasicsChk = 1
NSegs = Val(NSegsText)
RMax = Val(RMaxText)
ZMax = Val(ZMaxText)
RStart(1) = Val(RStartText)
ZStart(1) = Val(ZStartText)
SegNoMsg = "1"
SLOSHForm1.MultiPage1.Value = 1

End Sub




Private Sub DefSegOK_Click()
'
' Process the inputs from the Geometry Page
' Process the tank segments one by one
'
Dim i As Integer
Dim np_crv As Integer
Dim SegOK As Integer
'
' Make sure the Basic Page was successfully completed
If BasicsChk <> 1 Then
   MsgBox "You must complete the Basics Page first."
   SLOSHForm1.MultiPage1.Value = 0
   Exit Sub
End If
'
' Make Sure a segment type is chosen
If Not (StrSeg) Then
    If Not (CrcSeg) Then
        If Not (EllSeg) Then
            MsgBox "You must select a segment type."
            Exit Sub
        End If
    End If
End If
'
If (StrSeg) Then
' Make sure the Rend is entered and is within the envelope
    If REndStrText.Text = "" Then
        MsgBox "You must enter the R-coordinate of end of the segment."
        REndStrText.SetFocus
        Exit Sub
    End If
    If Val(REndStrText.Text) < 0# Or Val(REndStrText.Text) > RMax Then
        MsgBox "R-coordinate of segment end is outside tank envelope."
        REndStrText.SetFocus
        Exit Sub
    End If
' Make sure the Zend is entered and is within the envelope
    If ZEndStrText.Text = "" Then
        MsgBox "You must enter the Z-coordinate of end of the segment."
        ZEndStrText.SetFocus
        Exit Sub
    End If
    If Val(ZEndStrText.Text) < ZStart(1) Or Val(ZEndStrText.Text) > ZMax Then
        MsgBox "Z-coordinate of segment end is outside tank envelope."
        ZEndStrText.SetFocus
        Exit Sub
    End If
'
    SegType(ISeg1) = 1
    REnd(ISeg1) = Val(REndStrText)
    ZEnd(ISeg1) = Val(ZEndStrText)
    RO(ISeg1) = RStart(i)
    ZO(ISeg1) = ZStart(i)
    Rradius(ISeg1) = -999
    Zradius(ISeg1) = -999
    REndStrText = ""
    ZEndStrText = ""
End If
'
If (CrcSeg) Then
' Make sure the Rend is entered and is within the envelope
    If REndCrcText.Text = "" Then
        MsgBox "You must enter the R-coordinate of end of the segment."
        REndCrcText.SetFocus
        Exit Sub
    End If
    If Val(REndCrcText.Text) < 0 Or Val(REndCrcText.Text) > RMax Then
        MsgBox "R-coordinate of segment end is outside tank envelope."
        REndCrcText.SetFocus
        Exit Sub
    End If
' Make sure the Zend is entered and is within the envelope
    If ZEndCrcText.Text = "" Then
        MsgBox "You must enter the Z-coordinate of end of the segment."
        ZEndCrcText.SetFocus
        Exit Sub
    End If
    If Val(ZEndCrcText.Text) < ZStart(1) Or Val(ZEndCrcText.Text) > ZMax Then
        MsgBox "Z-coordinate of segment end is outside tank envelope."
        ZEndCrcText.SetFocus
        Exit Sub
    End If
' Make sure the Radius is entered
    If RadCrcText.Text = "" Then
        MsgBox "You must enter the Radius of the Circle."
        RadCrcText.SetFocus
        Exit Sub
    End If
' Make sure the RCtrCrc is entered
    If RCtrCrcText.Text = "" Then
        MsgBox "You must enter the R-coordinate of Circle Center."
        RCtrCrcText.SetFocus
        Exit Sub
    End If
' Make sure the ZCtrCrc is entered
    If ZCtrCrcText.Text = "" Then
        MsgBox "You must enter the Z_coordinate of Circle Center."
        ZCtrCrcText.SetFocus
        Exit Sub
    End If
' Make sure the segment coordinate pairs will be single-valued
    If (ZStart(ISeg1) - Val(ZCtrCrcText)) * (Val(ZEndCrcText) - Val(ZCtrCrcText)) < 0 Then
        MsgBox "Endpoints must not straddle the Circle Center."
        ZCtrCrcText.SetFocus
        Exit Sub
    End If
' Make sure the segment coordinate pairs will be single-valued
    If (RStart(ISeg1) - Val(RCtrCrcText)) * (Val(REndCrcText) - Val(RCtrCrcText)) < 0 Then
        MsgBox "Endpoints must not straddle the Circle Center."
        RCtrCrcText.SetFocus
        Exit Sub
    End If
'
    SegType(ISeg1) = 2
    REnd(ISeg1) = Val(REndCrcText)
    ZEnd(ISeg1) = Val(ZEndCrcText)
    RO(ISeg1) = Val(RCtrCrcText)
    ZO(ISeg1) = Val(ZCtrCrcText)
    Rradius(ISeg1) = Val(RadCrcText)
    Zradius(ISeg1) = Val(RadCrcText)
    REndCrcText = ""
    ZEndCrcText = ""
    RCtrCrcText = ""
    ZCtrCrcText = ""
    RadCrcText = ""
End If
'
'
If (EllSeg) Then
' Make sure the Rend is entered
    If REndEllText.Text = "" Then
        MsgBox "You must enter the R-coordinate of end of the segment."
        REndEllText.SetFocus
        Exit Sub
    End If
    If Val(REndEllText.Text) < 0 Or Val(REndEllText.Text) > RMax Then
        MsgBox "R-coordinate of segment end is outside tank envelope."
        REndEllText.SetFocus
        Exit Sub
    End If
' Make sure the Zend is entered and is within the envelope
    If ZEndEllText.Text = "" Then
        MsgBox "You must enter the Z-coordinate of end of the segment."
        ZEndEllText.SetFocus
        Exit Sub
    End If
    If Val(ZEndEllText.Text) < ZStart(1) Or Val(ZEndEllText.Text) > ZMax Then
        MsgBox "Z-coordinate of segment end is outside tank envelope."
        ZEndEllText.SetFocus
        Exit Sub
    End If
' Make sure the R-radius is entered
    If RRadEllText.Text = "" Then
        MsgBox "You must enter the Radius abaout R-axis of the Ellipse."
        RRadEllText.SetFocus
        Exit Sub
    End If
' Make sure the Z-radius is entered
    If ZRadEllText.Text = "" Then
        MsgBox "You must enter the Radius abaout Z-axis of the Ellipse."
        ZRadEllText.SetFocus
        Exit Sub
    End If
' Make sure the RCtrEll is entered
    If RCtrEllText.Text = "" Then
        MsgBox "You must enter the R-coordinate of Ellipse Center."
        RCtrEllText.SetFocus
        Exit Sub
    End If
' Make sure the ZCtrEll is entered
    If ZCtrEllText.Text = "" Then
        MsgBox "You must enter the Z-coordinate of Ellipse Center."
        ZCtrEllText.SetFocus
        Exit Sub
    End If
' Make sure the segment coordinate pairs will be single-valued
    If (ZStart(ISeg1) - Val(ZCtrEllText)) * (Val(ZEndEllText) - Val(ZCtrEllText)) < 0 Then
        MsgBox "Endpoints must not straddle the Ellipse Center."
        ZCtrEllText.SetFocus
        Exit Sub
    End If
' Make sure the segment coordinate pairs will be single-valued
    If (RStart(ISeg1) - Val(RCtrEllText)) * (Val(REndEllText) - Val(RCtrEllText)) < 0 Then
        MsgBox "Endpoints must not straddle the Ellipse Center."
        RCtrEllText.SetFocus
        Exit Sub
    End If
    SegType(ISeg1) = 3
    REnd(ISeg1) = Val(REndEllText)
    ZEnd(ISeg1) = Val(ZEndEllText)
    RO(ISeg1) = Val(RCtrEllText)
    ZO(ISeg1) = Val(ZCtrEllText)
    Rradius(ISeg1) = Val(RRadEllText)
    Zradius(ISeg1) = Val(ZRadEllText)
    REndEllText = ""
    ZEndEllText = ""
    RCtrEllText = ""
    ZCtrEllText = ""
    RRadEllText = ""
    ZRadEllText = ""
End If
'
' Draw Segment and ask User if it is OK
np_crv = 100
ReDim Seg1Poly(1 To np_crv, 1 To 2) As Single
Call DrawSeg(SegType(ISeg1), RStart(ISeg1), ZStart(ISeg1), REnd(ISeg1), ZEnd(ISeg1), RO(ISeg1), ZO(ISeg1), Rradius(ISeg1), Zradius(ISeg1), np_crv, Seg1Poly, Fname)
'
' Make the layout page visible and show current sketch
SLOSHForm1.MultiPage1.Pages(4).Visible = True
SLOSHForm1.TankShape.Width = PicWidth
SLOSHForm1.TankShape.Height = PicHeight
SLOSHForm1.TankShape.Visible = True
SLOSHForm1.TankShape.Picture = LoadPicture(Fname)
SLOSHForm1.MultiPage1.Value = 4
SegOK = MsgBox("Is the segment OK?", vbYesNo + vbQuestion + vbDefaultButton2)
'
' Hide the layout page
SLOSHForm1.TankShape.Visible = False
SLOSHForm1.MultiPage1.Pages(4).Visible = False
'
' If segment is NOT OK, return and get another set of coordinates
If SegOK = vbNo Then
   SLOSHForm1.MultiPage1.Value = 1
   REndStrText.SetFocus
   Exit Sub
End If
'
'  Segment is OK.  Add the current polyline to the aggregate polyline
For i = 1 To np_crv
   SegPoly(NppDraw + i, 1) = Seg1Poly(i, 1)
   SegPoly(NppDraw + i, 2) = Seg1Poly(i, 2)
Next i
NppDraw = NppDraw + np_crv
ISeg1 = ISeg1 + 1
'
'  If this is the last segment, process the segment geometry into the SLOSH code geometry variables
'
If ISeg1 > NSegs Then
   For i = 1 To NSegs
      If SegType(i) = 1 Then
'
' 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1
' A1=A3=A5=0 for straight wall segment
' A2=1
' A4=negative of slope
         A1(i) = 0
         A3(i) = 0
         A5(i) = 0
         A2(i) = 1
         RO(i) = RStart(i)
         ZO(i) = ZStart(i)
'
' Don't allow exactly zero or infinite slope for straight segments
         If Abs(ZEnd(i) - ZStart(i)) < 0.000001 Then
            A4(i) = -0.000000001
          ElseIf Abs(REnd(i) - RStart(i)) < 0.000001 Then
            A4(i) = -100000000#
           Else
            A4(i) = -(ZEnd(i) - ZStart(i)) / (REnd(i) - RStart(i))
         End If
      End If
' 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2
' Circular arc wall segment. Center can be off the tank axis
' A1*(z-Zc)^2 + A3*(r-Rc)^2 + A5 = 0
      If SegType(i) >= 2 Then
         If SegType(i) = 2 Then
            Zradius(i) = Rradius(i)
         End If
'
' A2=A4=0.
' A1=1 A3=(Zr/Rr)^2, and A5 = -(Zr)^2
         A2(i) = 0
         A4(i) = 0
         A1(i) = 1
         A3(i) = (Zradius(i) / Rradius(i)) ^ 2
         A5(i) = -(Zradius(i)) ^ 2
      End If
' 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2
   Next i
'
' All geometry has been processed
' Go to Fluid Properties Page
'
   GeomDefChk = 1
   SLOSHForm1.MultiPage1.Value = 2
   Exit Sub
End If
'
'  This was not the last segment --- return and get the next segment
If ISeg1 >= 2 Then
   RStart(ISeg1) = REnd(ISeg1 - 1)
   ZStart(ISeg1) = ZEnd(ISeg1 - 1)
End If
'
SegNoMsg = Str(ISeg1)
StrSeg = False
EllSeg = False
CrcSeg = False
'
SLOSHForm1.MultiPage1.Value = 1
'
End Sub



Private Sub PropsOK_Click()
'
' Process the inputs from the Fluid Properties Page
'
Dim ZMin As Integer
Dim temp As Double
Dim LevOK  As Integer
'
' Make sure the Basics Page and Geometry were successfully completed.
'
PropDefChk = 0
If BasicsChk <> 1 Then
   MsgBox "You must complete the Basics Page first."
   SLOSHForm1.MultiPage1.Value = 0
   Exit Sub
End If
If GeomDefChk <> 1 Then
   MsgBox "You must complete the Geometry Page first."
   SLOSHForm1.MultiPage1.Value = 1
   Exit Sub
End If
'
' Make sure the Density is entered
If DensText.Text = "" Then
    MsgBox "You must enter the Density."
    DensText.SetFocus
    Exit Sub
End If
'
' Make sure the Viscosity is entered
If ViscText.Text = "" Then
    MsgBox "You must enter the Kin. Viscosity."
    ViscText.SetFocus
    Exit Sub
End If
'
' Make sure the Gravity is entered
If GravText.Text = "" Then
    MsgBox "You must enter the Gravity(Thrust)."
    GravText.SetFocus
    Exit Sub
End If
'
' Make sure the Liquid Level is entered
If ZLiquidText.Text = "" Then
    MsgBox "You must enter the Liquid Surface Z-coordinate."
    ZLiquidText.SetFocus
    Exit Sub
End If
'
'Make sure the Liquid Level is within the tank axial limits
If Val(ZLiquidText.Text) >= ZMax Or Val(ZLiquidText.Text) <= ZMin Then
    MsgBox "Liquid Surface Z-coordinate must be within the tank."
    ZLiquidText.SetFocus
    Exit Sub
End If
'
' Find the radial extents of the free surface for plotting
'
ZLiquid = Val(ZLiquidText)
temp = FreeSurf()
'
'  Draw the liqiud level for the user to review
'
temp = DrawLev(ZLiquid)
'
' Make the layout page visible and show current sketch
SLOSHForm1.MultiPage1.Pages(4).Visible = True
SLOSHForm1.TankShape.Width = PicWidth
SLOSHForm1.TankShape.Height = PicHeight + 20
SLOSHForm1.TankShape.Visible = True
SLOSHForm1.TankShape.Picture = LoadPicture(Fname)
SLOSHForm1.MultiPage1.Value = 4
LevOK = MsgBox("Is the Liquid Level OK?", vbYesNo + vbQuestion + vbDefaultButton2)
'
' Hide the layout page
SLOSHForm1.TankShape.Visible = False
SLOSHForm1.MultiPage1.Pages(4).Visible = False
'
'  Level is not OK  --- return and get new value
If LevOK = vbNo Then
   SLOSHForm1.MultiPage1.Value = 2
   ZLiquidText.SetFocus
   Exit Sub
End If
'
' Level is OK
' Transfer numerical values and go to Damping Page
'
Density = Val(DensText)
Viscosity = Val(ViscText)
GLevel = Val(GravText)
PropDefChk = 1
'
SLOSHForm1.MultiPage1.Value = 3
'
End Sub

Private Sub DampOK_Click()
'
' Process the inputs from the Damping Model Page
'
Dim ZL As Double
Dim temp As Integer
Dim SLOSHOK As Integer
'
' Make sure the Basics PAge, Geomtetry Page, and FLuid Properties Page were successfully completed.
'
DampDefChk = 0
If BasicsChk <> 1 Then
   MsgBox "You must complete the Basics Page first."
   SLOSHForm1.MultiPage1.Value = 0
   Exit Sub
End If
If GeomDefChk <> 1 Then
   MsgBox "You must complete the Geometry Page first."
   SLOSHForm1.MultiPage1.Value = 1
   Exit Sub
End If
If PropDefChk <> 1 Then
   MsgBox "You must complete the Fluid Properties Page first."
   SLOSHForm1.MultiPage1.Value = 2
   Exit Sub
End If
'
' Make Sure a damping model is chosen is
If Not (DampAnn) Then
   If Not (DampCyl) Then
      If Not (DampSpr) Then
         If Not (DampSpr) Then
            If Not (DampCon) Then
               If Not (DampTor) Then
                  If Not (DampBaf) Then
                     MsgBox "You must select a damping model type."
                     SLOSHForm1.MultiPage1.Value = 3
                     Exit Sub
                  End If
               End If
            End If
         End If
      End If
   End If
End If
'
' Initialize parameters if this is a brand new problem setup.
'
If Not (ReadSheet) Then
   IndexType_damp = -999
   ZL_damp = -999
   RTank_damp = -999
   RIn_damp = -999
   Angle_damp = -999
   BW_damp = -999
   Axspc_damp = -999
   NumSub_damp = -999
   TopSub_damp = -999
   IndexType_damp = -999
   NumSub_damp = -999
   IndexBaffle = -999
End If
'
If (DampAnn) Then
' Make sure the Fill Level for Damping model is entered
    If ZFillDampAnnText.Text = "" Then
        MsgBox "You must enter the fill level for the damping model."
        ZFillDampAnnText.SetFocus
        Exit Sub
    End If
' Make sure the Radii are entered
    If RTankDampAnnText.Text = "" Then
        MsgBox "You must enter the outer radius of tank at surface."
        RTankDampAnnText.SetFocus
        Exit Sub
    End If
    If RTankInAnnText.Text = "" Then
        MsgBox "You must enter the inner radius of tank at surface."
        RTankInAnnText.SetFocus
        Exit Sub
    End If
    IndexType_damp = 1
    ZL_damp = Val(ZFillDampAnnText)
    RTank_damp = Val(RTankDampAnnText)
    RIn_damp = Val(RTankInAnnText)
End If
'
If (DampCyl) Then
' Make sure the Fill Level for Damping model is entered
    If ZFillDampCylText.Text = "" Then
        MsgBox "You must enter the fill Z-level for the damping model."
        ZFillDampCylText.SetFocus
        Exit Sub
    End If
' Make sure the Radii are entered
    If RTankDampCylText.Text = "" Then
        MsgBox "You must enter the outer radius of tank at surface."
        RTankDampCylText.SetFocus
        Exit Sub
    End If
    IndexType_damp = 2
    ZL_damp = Val(ZFillDampCylText)
    RTank_damp = Val(RTankDampCylText)
End If
'
If (DampSpr) Then
' Make sure the Fill Level for Damping model is entered
    If ZFillDampSprText.Text = "" Then
        MsgBox "You must enter the fill Z-level for the damping model."
        ZFillDampSprText.SetFocus
        Exit Sub
    End If
' Make sure the Radii are entered
    If RTankDampSprText.Text = "" Then
        MsgBox "You must enter the outer radius of tank at surface."
        RTankDampSprText.SetFocus
        Exit Sub
    End If
    IndexType_damp = 3
    ZL_damp = Val(ZFillDampSprText)
    RTank_damp = Val(RTankDampSprText)
End If
'
If (DampTor) Then
' Make sure the Fill Level for Damping model is entered
    If ZFillDampTorText.Text = "" Then
        MsgBox "You must enter the fill Z-level for the damping model."
        ZFillDampTorText.SetFocus
        Exit Sub
    End If
' Make sure the Radii are entered
    If RTankDampTorText.Text = "" Then
        MsgBox "You must enter the outer radius of tank at surface."
        RTankDampTorText.SetFocus
        Exit Sub
    End If
    IndexType_damp = 4
    ZL_damp = Val(ZFillDampTorText)
    RTank_damp = Val(RTankDampTorText)
End If
'
If (DampCon) Then
' Make sure the Fill Level for Damping model is entered
    If ZFillDampConText.Text = "" Then
        MsgBox "You must enter the fill Z-level for the damping model."
        ZFillDampConText.SetFocus
        Exit Sub
    End If
' Make sure the Angle is entered
    If AngDampConText.Text = "" Then
        MsgBox "You must enter the tank wall angle."
        AngDampConText.SetFocus
        Exit Sub
    End If
' Make sure a positive angle is entered
    If Val(AngDampConText.Text) < 0 Then
        MsgBox "You must enter a positive angle."
        AngDampConText = ""
        AngDampConText.SetFocus
        Exit Sub
    End If
    IndexType_damp = 5
    ZL_damp = Val(ZFillDampConText)
    Angle_damp = Val(AngDampConText) * PiVal / 180#
    RTank_damp = ZL_damp * Tan(Angle_damp)
End If
'
If (DampBaf) Then
' Make sure the Fill Level for Damping model is entered
    If ZFillDampBafText.Text = "" Then
        MsgBox "You must enter the fill Z-level for the damping model."
        ZFillDampBafText.SetFocus
        Exit Sub
    End If
' Make sure the Radii are entered
    If RTankDampBafText.Text = "" Then
        MsgBox "You must enter the outer radius of tank at surface."
        RTankDampBafText.SetFocus
        Exit Sub
    End If
' Make sure the Baffle Width is entered
    If BWidDampBafText.Text = "" Then
        MsgBox "You must enter the Baffle Width."
        BWidDampBafText.SetFocus
        Exit Sub
    End If
' Make sure the Axial Spacing is entered
    If ASpcDampBafText.Text = "" Then
        MsgBox "You must enter the Baffle Width."
        ASpcDampBafText.SetFocus
        Exit Sub
    End If
' Make sure the Number of Submerged Baffles is entered
    If NSubDampBafText.Text = "" Then
        MsgBox "You must enter the Number of Submerged Baffles."
        NSubDampBafText.SetFocus
        Exit Sub
    End If
' Make sure the Submerged Depth of Top Baffle is entered
    If TSubDampBafText.Text = "" Then
        MsgBox "You must enter the Submerged Depth of Top Baffle."
        TSubDampBafText.SetFocus
        Exit Sub
    End If
'
    IndexType_damp = 6
    ZL_damp = Val(ZFillDampBafText)
    RTank_damp = Val(RTankDampBafText)
    BW_damp = Val(BWidDampBafText)
    Axspc_damp = Val(ASpcDampBafText)
    NumSub_damp = Val(NSubDampBafText)
    TopSub_damp = Val(TSubDampBafText)
    IndexBaffle = 1
End If
'
' All inputs seem OK ---- Record the entire problem setup in a worksheet,
'
temp = WriteInputs()
'
'  Make sure user wants to execute the SLOSH code
'
SLOSHOK = MsgBox("All Inputs Entered, OK to Execute SLOSH?", vbYesNo + vbQuestion + vbDefaultButton2)
If SLOSHOK = vbYes Then
   Call SloshExec
End If
   Application.ScreenUpdating = True
   SLOSHForm1.Hide
'   Exit Sub

End Sub



Sub SloshExec()
'
'  Excel version SLOSH with calculations performed in a DLL.
'  The usage of the SLOSH
'
Dim nsegmx_pass As Long
Dim IAi1(1 To 5) As Long
Dim IAo1(1 To 1) As Long
Dim RAi1(1 To 13) As Double
Dim RAo1(1 To 15) As Double
Dim RAo2(1 To 13, 1 To 7) As Double
Dim IAi2(1 To 20) As Long
Dim RAi2(1 To 20, 1 To 8) As Double
Dim temp As Variant
Dim i As Integer
Dim NW_1  As Long
Dim H1 As Double
Dim H2 As Double
Dim PL1 As Double
Dim PL2 As Double
Dim PM1 As Double
Dim PM2 As Double
Dim PM0 As Double
Dim H0 As Double
Dim FrozI As Double
Dim Rat1 As Double
Dim Rat2 As Double
Dim Zeta As Double
Dim TotalMass As Double
Dim ZL_sl As Double
Dim AvPress As Double

For i = 1 To 13
   ZP_1(i) = 0
   RPOut_1(i) = 0
   POut1_1(i) = 0
   POut2_1(i) = 0
   RPIn_1(i) = 0
   PIn1_1(i) = 0
   PIn2_1(i) = 0
Next i
'
' Prepare for the call to SLOSH_CALC
' All arguments must be passed in arrays.
'
nsegmx_pass = 20
IAi1(1) = nsegmx_pass
IAi1(2) = NSegs
IAi1(3) = IndexType_damp
IAi1(4) = NumSub_damp
IAi1(5) = IndexBaffle
'
RAi1(1) = RMax
RAi1(2) = ZMax
RAi1(3) = ZLiquid
RAi1(4) = Density
RAi1(5) = Viscosity
RAi1(6) = GLevel
RAi1(7) = ZL_damp
RAi1(8) = RTank_damp
RAi1(9) = RIn_damp
RAi1(10) = Angle_damp
RAi1(11) = BW_damp
RAi1(12) = Axspc_damp
RAi1(13) = TopSub_damp
'
For i = 1 To NSegs
   IAi2(i) = SegType(i)
   RAi2(i, 1) = RStart(i)
   RAi2(i, 2) = ZStart(i)
   RAi2(i, 3) = REnd(i)
   RAi2(i, 4) = ZEnd(i)
   RAi2(i, 5) = RO(i)
   RAi2(i, 6) = ZO(i)
   RAi2(i, 7) = Rradius(i)
   RAi2(i, 8) = Zradius(i)
Next i
'
'  Call to SLOSH_CALC for slosh model parameters.
'
Call SLOSH_Calc(IAi1(1), IAi2(1), IAo1(1), RAi1(1), RAi2(1, 1), RAo1(1), RAo2(1, 1))
'
'  Transfer SLOSH_CALC outputs to varaible names that make sense.
'
NW_1 = IAo1(1)
NInner = 0
If NW_1 < 0 Then NInner = 1
NW_1 = Abs(NW_1)
'
H1 = RAo1(1)
H2 = RAo1(2)
PL1 = RAo1(3)
PL2 = RAo1(4)
PM1 = RAo1(5)
PM2 = RAo1(6)
PM0 = RAo1(7)
H0 = RAo1(8)
FrozI = RAo1(9)
Rat1 = RAo1(10)
Rat2 = RAo1(11)
Zeta = RAo1(12)
TotalMass = RAo1(13)
ZL_sl = RAo1(14)
AvPress = RAo1(15)
For i = 1 To NW_1
   ZP_1(i) = RAo2(i, 1)
   RPOut_1(i) = RAo2(i, 2)
   POut1_1(i) = RAo2(i, 3)
   POut2_1(i) = RAo2(i, 4)
   RPIn_1(i) = RAo2(i, 5)
   PIn1_1(i) = RAo2(i, 6)
   PIn2_1(i) = RAo2(i, 7)
Next i
'
'  Transfer outputs to worksheet   TANK.OUT
'
temp = OutputModel(NInner, NW_1, ZLiquid, H1, H2, PL1, PL2, PM1, PM2, Rat1, Rat2, PM0, H0, FrozI, Zeta, TotalMass, IndexBaffle, AvPress)
End Sub



Function OutputModel(NInner, NW_1, LHt, Ht1, Ht2, PL1, PL2, PM1, PM2, Wv1, Wv2, FixM, FixH, FixI, Zeta, Mas, IndexBaffle, Avp)
'
'  Writes all SLOSH code outputs to the worksheet
'
Dim iw As Integer
Dim Text As String
'
'
Application.ScreenUpdating = False
Range(Cells(LastrowIn + 2, 1), Cells(LastrowIn + 2, 1)).Select
ActiveCell.Value = "OUTPUTS"
With Selection.Font
    .Name = "Calibri"
    .Size = 14
End With
With Selection.Font
    .Color = -16776961
    .TintAndShade = 0
End With
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 65535
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

Sheets(SheetName.Text).Select
Range(Cells(LastrowIn + 3, 1), Cells(LastrowIn + 3, 1)).Select
Selection.ColumnWidth = 45
Selection.NumberFormat = "0.000E+00"
Range(Cells(LastrowIn + 3, 1), Cells(LastrowIn + 3, 1)).Activate
Range(Cells(LastrowIn + 3, 1), Cells(LastrowIn + 3, 1)).Activate
'
'  Total liquid mass and liquid surface height
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "LIQUID MASS [mass units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = Mas
Text = Format(Mas, "0.000E+00")
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "LIQUID SURFACE HEIGHT above z=0 [length units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = LHt
'
' First slosh mode values
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "FIRST MODE PARAMETERS"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "Pendulum mass [mass units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = PM1
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "Pendulum length [length units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = PL1
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "Pendulum hinge z-location [length units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = Ht1
'
ActiveCell.Offset(1, -1).Activate
If IndexBaffle = 1 Then
    ActiveCell.Value = "BAFFLE Pend. % crit. damp./(Slosh amp./Tank Rad)^.5"
  Else
    ActiveCell.Value = "Pendulum % critical damping"
End If
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = Zeta
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "Ratio of slosh amplitude to pendulum amplitude"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = Wv1
'
' Second slosh mode values
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "SECOND MODE PARAMETERS"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "Pendulum mass [mass units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = PM2
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "Pendulum length [length units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = PL2
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "Pendulum hinge z-location [length units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = Ht2
'
ActiveCell.Offset(1, -1).Activate
If IndexBaffle = 1 Then
    ActiveCell.Value = "BAFFLE Pend. % crit. damp./(Slosh amp./Tank Rad)^.5"
   Else
    ActiveCell.Value = "Pendulum % critical damping"
End If
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = Zeta
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "Ratio of slosh amplitude to pendulum amplitude"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = Wv2
'
' Fixed mass parameters
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "FIXED MASS PARAMETERS"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "Mass [mass units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = FixM
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "Z-location [length units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = FixH
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "Mom. Inertia [mass*length^2 units]"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = FixI
'
'  Record the Damping Model description
ActiveCell.Offset(2, -1).Activate
ActiveCell.Value = "Damping Model"
ActiveCell.Offset(0, 1).Activate
Select Case IndexType_damp
'
Case 1                                                  '       annular tank
    ActiveCell.Value = "Annular Tank"
Case 2                                                  '       cylindrical tank
    ActiveCell.Value = "Cylindrical Tank"
Case 3                                                  '       spherical tank
    ActiveCell.Value = "Spherical Tank"
Case 4                                                  '       toroidal tank
    ActiveCell.Value = "Toroidal Tank"
Case 5                                                  '       conical tank
    ActiveCell.Value = "Conical Tank"
Case 6                                                  '       baffled cylindrical tank
    ActiveCell.Value = "Baffled Cylindrical Tank"
End Select
'
' Wall Pressure Coefficients
ActiveCell.Offset(2, -1).Activate
ActiveCell.Value = "PRESSURE COEFFICIENT AMPLITUDE AT WALLS"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "(Coefficients = amplitude of the cosine wave around the tank"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "Units are force/length^2 per unit 1st mode mass lateral deflection.)"
ActiveCell.Offset(2, 0).Activate
ActiveCell.Value = "Z"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "Outer_R"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "1st MODE"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "2nd MODE"
'If NInner > 0 Then
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "Inner_R"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "1st MODE"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "2nd MODE"
ActiveCell.Offset(0, -3).Activate
'End If
ActiveCell.Offset(1, -3).Activate
For iw = 1 To NW_1
    ActiveCell.Value = ZP_1(iw)
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Value = RPOut_1(iw)
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Value = POut1_1(iw)
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.Value = POut2_1(iw)
    If (NInner > 0) And (iw <> NW_1) Then
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = RPIn_1(iw)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = PIn1_1(iw)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = PIn2_1(iw)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Offset(0, -4).Activate
    Else
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "---"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "---"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "---"
        ActiveCell.Offset(0, -3).Activate
    End If
    ActiveCell.Offset(1, -3).Activate
Next iw
ActiveCell.Offset(2, 0).Activate
'
'  WIDTH-AVERAGE BAFFLE PRESSURE (for baffle at liquid surface)
If IndexBaffle > 0 Then
    ActiveCell.Value = "WIDTH-AVERAGE BAFFLE PRESSURE (for baffle at liquid surface):"
    ActiveCell.Offset(1, 0).Activate
    ActiveCell.Value = "force/length ^2 per unit 1st mode mass lateral deflection"
    ActiveCell.Offset(1, 0).Activate
    ActiveCell.Value = Avp
End If
Range(Cells(LastrowIn + 2, 1), Cells(LastrowIn + 2, 1)).Activate
Application.ScreenUpdating = True
'
'
End Function




Function WriteInputs()
'
'  Records all the input paramters so that entire problem setup can be re-read later
'
Dim i As Integer
Dim FinalPic As Object
'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'
' Create the worksheet.  This may be double duty, but earlier creations
' were needed to keep the screen from flashing between worksheets while
' the tank inputs were being processed.
If ((SheetExists(SheetName.Text))) Then
   Sheets(SheetName.Text).Delete
End If
Sheets.Add
ActiveSheet.Select
ActiveSheet.Name = SheetName.Text
ActiveSheet.Move after:=Worksheets(Worksheets.Count)
ActiveSheet.Tab.ColorIndex = xlColorIndexNone
Range("A1").Select

ActiveCell.Value = "INPUTS"
With Selection.Font
    .Name = "Calibri"
    .Size = 14
End With
With Selection.Font
    .Color = -16776961
    .TintAndShade = 0
End With
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 65535
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
Range("A2").Select
ActiveCell.Value = Comment.Text
'
'  Rmax Zmax
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "Rmax"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "Zmax"
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = RMax
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = ZMax
'
'  NSegs
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = "NSegs"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = NSegs
'
'  RStart   ZStart
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "RStart"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "ZStart"
ActiveCell.Offset(1, -1).Activate
ActiveCell.Value = RStart(1)
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = ZStart(1)
'
'  Segments
ActiveCell.Offset(1, -1).Activate
For i = 1 To NSegs
    ActiveCell.Value = "Segment " & Format(i, "0") & " Type"
    ActiveCell.Offset(1, 0).Activate
    ActiveCell.Value = SegType(i)
    ActiveCell.Offset(1, 0).Activate
    If SegType(i) = 1 Then
        ActiveCell.Value = "REnd"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "ZEnd"
        ActiveCell.Offset(1, -1).Activate
        ActiveCell.Value = REnd(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = ZEnd(i)
        ActiveCell.Offset(1, -1).Activate
    End If
    If SegType(i) = 2 Then
        ActiveCell.Value = "REnd"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "ZEnd"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "RCenter"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "ZCenter"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "Radius"
        ActiveCell.Offset(1, -4).Activate
        ActiveCell.Value = REnd(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = ZEnd(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = RO(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = ZO(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = Rradius(i)
        ActiveCell.Offset(1, -4).Activate
    End If
    If SegType(i) = 3 Then
        ActiveCell.Value = "REnd"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "ZEnd"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "RCenter"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "ZCenter"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "RRadius"
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "ZRadius"
        ActiveCell.Offset(1, -5).Activate
        ActiveCell.Value = REnd(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = ZEnd(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = RO(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = ZO(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = Rradius(i)
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = Zradius(i)
        ActiveCell.Offset(1, -5).Activate
    End If
Next i
'
'  Properties
ActiveCell.Value = "Liq.Ht"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = ZLiquid
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "Density"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "Kin.Viscosity"
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = "Gravity"
ActiveCell.Offset(1, -2).Activate
ActiveCell.Value = Density
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = Viscosity
ActiveCell.Offset(0, 1).Activate
ActiveCell.Value = GLevel
ActiveCell.Offset(1, -2).Activate
'
'   Damping
ActiveCell.Value = "Damp_Model 1=Annular, 2=Cylinder, 3=Sphere, 4=Toroid, 5=Cone, 6=Baffled Cylinder"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = IndexType_damp
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = "Z.Liquid.Damp"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = ZL_damp
ActiveCell.Offset(-1, 1).Activate
ActiveCell.Value = "ROuter.Damp"
ActiveCell.Offset(1, 0).Activate
ActiveCell.Value = RTank_damp
If IndexType_damp = 1 Then
   ActiveCell.Offset(-1, 1).Activate
   ActiveCell.Value = "RInner.Damp"
   ActiveCell.Offset(1, 0).Activate
    ActiveCell.Value = RIn_damp
End If
If IndexType_damp = 5 Then
   ActiveCell.Offset(-1, 0).Activate
   ActiveCell.Value = "Angle.Damp.deg"
   ActiveCell.Offset(1, 0).Activate
   ActiveCell.Value = Angle_damp * 180 / PiVal
End If
If IndexType_damp = 6 Then
   ActiveCell.Offset(-1, 1).Activate
   ActiveCell.Value = "Baffle.Width"
   ActiveCell.Offset(1, 0).Activate
   ActiveCell.Value = BW_damp
   ActiveCell.Offset(-1, 1).Activate
   ActiveCell.Value = "Baffle.Axial.Space"
   ActiveCell.Offset(1, 0).Activate
   ActiveCell.Value = Axspc_damp
   ActiveCell.Offset(-1, 1).Activate
   ActiveCell.Value = "Num.Baf.Submerged"
   ActiveCell.Offset(1, 0).Activate
   ActiveCell.Value = NumSub_damp
   ActiveCell.Offset(-1, 1).Activate
   ActiveCell.Value = "Top.Baf.Depth"
   ActiveCell.Offset(1, 0).Activate
   ActiveCell.Value = TopSub_damp
End If
'
' Add the tank sketch to the worksheet
PicX0 = 450
PicZ0 = 25
Set FinalPic = ActiveSheet.Shapes.AddPicture(Fname, linktofile:=msoFalse, _
    savewithdocument:=msoCTrue, Left:=PicX0, Top:=PicZ0, Width:=PicWidth, Height:=PicHeight + 20)
FinalPic.Name = SheetName.Text & "_sketch"
'
' Delete the picture file if it still exists
If Dir(Fname) <> "" Then
' First remove readonly attribute, if set
   SetAttr Fname, vbNormal
' Then delete the file
   Kill Fname
End If
'
' Find the last row of the Inputs
Range("A1").Select
Selection.End(xlDown).Select
LastrowIn = ActiveCell.Row
'Application.ScreenUpdating = True

   
End Function




Function ReadTank(SheetName)
'
'  Reads a problem setup from a previously executed session
'
Dim i As Integer
Dim Damp_Ann As Boolean
Dim Damp_Cyl As Boolean
Dim Damp_Spr As Boolean
Dim Damp_Tor As Boolean
Dim Damp_Con As Boolean
Dim Damp_Baf As Boolean
Dim ZL As Double
Dim angle As Double
Dim i1 As Integer
Dim np_crv As Integer
Dim i2 As Integer

'
Sheets(SheetName.Text).Select
Sheets(SheetName.Text).Activate
Application.ScreenUpdating = False
Application.DisplayAlerts = True
'
' Process all inputs into variables values.
'
Range("a2").Select
Comment.Text = ActiveCell.Value
'
' RMax, ZMax
Range("a4").Select
RMax = ActiveCell.Value
'
ActiveCell.Offset(0, 1).Activate
ZMax = ActiveCell.Value
'
' NSegs
Range("a6").Select
NSegs = ActiveCell.Value
'
' RStart, ZStart
Range("a8").Select
RStart(1) = ActiveCell.Value
'
ActiveCell.Offset(0, 1).Activate
ZStart(1) = ActiveCell.Value
'
' Segments
Range("a8").Select
For i = 1 To NSegs
   If i > 1 Then
      RStart(i) = REnd(i - 1)
      ZStart(i) = ZEnd(i - 1)
   End If
   ActiveCell.Offset(2, 0).Activate
   SegType(i) = ActiveCell.Value
' 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1
' Straight wall segment. Can be vertical, horizontal or cone wall
' A2*(z-Zstart)+A4*(r-Rstart)=0
   If SegType(i) = 1 Then
      ActiveCell.Offset(2, 0).Activate
      REnd(i) = ActiveCell.Value
      ActiveCell.Offset(0, 1).Activate
      ZEnd(i) = ActiveCell.Value
      ActiveCell.Offset(0, -1).Activate
'
' A1=A3=A5=0 for straight wall segment
' A2=1
' A4=negative of slope
      A1(i) = 0
      A3(i) = 0
      A5(i) = 0
      A2(i) = 1
      RO(i) = RStart(i)
      ZO(i) = ZStart(i)
'
' Don't allow exactly zero or infinite slope for straight segments
      If Abs(ZEnd(i) - ZStart(i)) < 0.000001 Then
          A4(i) = -0.000000001
        ElseIf Abs(REnd(i) - RStart(i)) < 0.000001 Then
          A4(i) = -100000000#
        Else
          A4(i) = -(ZEnd(i) - ZStart(i)) / (REnd(i) - RStart(i))
      End If
   End If
' 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1
' 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2
' Circular arc wall segment. Center can be off the tank axis
' A1*(z-Zc)^2 + A3*(r-Rc)^2 + A5 = 0
    If SegType(i) >= 2 Then
       ActiveCell.Offset(2, 0).Activate
       REnd(i) = ActiveCell.Value
       ActiveCell.Offset(0, 1).Activate
       ZEnd(i) = ActiveCell.Value
       ActiveCell.Offset(0, 1).Activate
       RO(i) = ActiveCell.Value
       ActiveCell.Offset(0, 1).Activate
       ZO(i) = ActiveCell.Value
       ActiveCell.Offset(0, 1).Activate
       Rradius(i) = ActiveCell.Value
       If SegType(i) = 2 Then
           Zradius(i) = Rradius(i)
         Else
           ActiveCell.Offset(0, 1).Activate
           Zradius(i) = ActiveCell.Value
           ActiveCell.Offset(0, -1).Activate
       End If
       ActiveCell.Offset(0, -4).Activate

'  ERROR CHECKS HERE
'  Rstart, Rend, Zstart, Zend outside of box
'  Rstart, Rend, Zstart, Zend not sufficiently close to equation
'  Rr = 0
'
' A2=A4=0.
' A1=1 A3=(Zr/Rr)^2, and A5 = -(Zr)^2
       A2(i) = 0
       A4(i) = 0
       A1(i) = 1
       A3(i) = (Zradius(i) / Rradius(i)) ^ 2
       A5(i) = -(Zradius(i)) ^ 2
    End If
' 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2 2
Next i
'
'  Fluid Properties and Fluid Level
'
ActiveCell.Offset(2, 0).Activate
ZLiquid = ActiveCell.Value
ActiveCell.Offset(2, 0).Activate
Density = ActiveCell.Value
ActiveCell.Offset(0, 1).Activate
Viscosity = ActiveCell.Value
ActiveCell.Offset(0, 1).Activate
GLevel = ActiveCell.Value
'
' Damping Model Inputs
ActiveCell.Offset(2, -2).Activate
IndexType_damp = ActiveCell.Value
IndexBaffle = 0
'
'  If Case <1 or > 6 Error
' Initialize all values
Damp_Ann = False
Damp_Cyl = False
Damp_Spr = False
Damp_Tor = False
Damp_Con = False
Damp_Baf = False
ZL_damp = -999
RTank_damp = -999
RIn_damp = -999
Angle_damp = -999
BW_damp = -999
Axspc_damp = -999
NumSub_damp = -999
TopSub_damp = -999
IndexBaffle = -999
'
Select Case IndexType_damp
Case 1                                                  '       annular tank
   DampAnn = True
   ActiveCell.Offset(2, 0).Activate
   ZL_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   RTank_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   RIn_damp = ActiveCell.Value
    
Case 2                                                      '       cylindrical tank
   DampCyl = True
   ActiveCell.Offset(2, 0).Activate
   ZL_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   RTank_damp = ActiveCell.Value

Case 3                                                      '       spherical tank
   DampSpr = True
   ActiveCell.Offset(2, 0).Activate
   ZL_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   RTank_damp = ActiveCell.Value

Case 4                                                          '       toroidal tank
   DampTor = True
   ActiveCell.Offset(2, 0).Activate
   ZL_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   RTank_damp = ActiveCell.Value

Case 5                                                              '       conical tank
   DampCon = True
   ActiveCell.Offset(2, 0).Activate
   ZL_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   Angle_damp = ActiveCell.Value * PiVal / 180#
   RTank_damp = ZL * Tan(Angle_damp)
Case 6                                                              '       baffled cylindrical tank
   DampBaf = True
   ActiveCell.Offset(2, 0).Activate
   ZL_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   RTank_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   BW_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   Axspc_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   NumSub_damp = ActiveCell.Value
   ActiveCell.Offset(0, 1).Activate
   TopSub_damp = ActiveCell.Value
   IndexBaffle = 1
End Select
   
' Regenerate the Picture
'  Call DrawBox to get the scaling and the envelope coordinates for the polyline
Fname = DrawBox(RMax, ZMax)
'
'  Call DrawSeg for each tank segment
NppDraw = 0
For i1 = 1 To NSegs
   np_crv = 100
   ReDim Seg1Poly(1 To np_crv, 1 To 2) As Single
   Call DrawSeg(SegType(i1), RStart(i1), ZStart(i1), REnd(i1), ZEnd(i1), RO(i1), ZO(i1), Rradius(i1), Zradius(i1), np_crv, Seg1Poly, Fname)
'Add the completed polyline to the total polyline
   For i2 = 1 To np_crv
      SegPoly(NppDraw + i2, 1) = Seg1Poly(i2, 1)
      SegPoly(NppDraw + i2, 2) = Seg1Poly(i2, 2)
   Next i2
   NppDraw = NppDraw + np_crv
Next i1
'
Application.ScreenUpdating = True
    
End Function




Function DrawBox(RMax, ZMax)
'
'  Draws the Tank Envelope
'
Dim SLOSH_Sheet As Worksheet
Dim XChart As Shape
Dim xxx As Shape
Dim filtname As Variant
'
'filtname = "JPG"
filtname = "BMP"
'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
If ((SheetExists("SLOSH_GRAPH"))) Then
   Sheets("SLOSH_GRAPH").Delete
End If
Set SLOSH_Sheet = Sheets.Add
SLOSH_Sheet.Select
SLOSH_Sheet.Name = "SLOSH_GRAPH"
Sheets("SLOSH_GRAPH").Move after:=Worksheets(Worksheets.Count)
Range("A1").Select
'
' Drawing coordinate transformation paramters
'
x0_draw = 50
z0_draw = 50
xzmx_draw = 300
scale_draw = xzmx_draw / WorksheetFunction.Max(RMax, ZMax)
zmx_draw = scale_draw * ZMax
'
' If this is the first pass, set the picture width and height
If BasicsOK = 0 Then
'   PicWidth = Val(RMaxText) * scale_draw * 1.05
'   PicHeight = Val(ZMaxText) * scale_draw * 1.05
   PicWidth = RMax * scale_draw * 1.05
   PicHeight = ZMax * scale_draw * 1.05
End If
'
' Draw the box for the tank envelope
'
EnvPoly(1, 1) = x_draw(0, x0_draw, scale_draw)
EnvPoly(1, 2) = z_draw(0, z0_draw, zmx_draw, scale_draw)
EnvPoly(2, 1) = x_draw(RMax, x0_draw, scale_draw)
EnvPoly(2, 2) = z_draw(0, z0_draw, zmx_draw, scale_draw)
EnvPoly(3, 1) = x_draw(RMax, x0_draw, scale_draw)
EnvPoly(3, 2) = z_draw(ZMax, z0_draw, zmx_draw, scale_draw)
EnvPoly(4, 1) = x_draw(0, x0_draw, scale_draw)
EnvPoly(4, 2) = z_draw(ZMax, z0_draw, zmx_draw, scale_draw)
EnvPoly(5, 1) = x_draw(0, x0_draw, scale_draw)
EnvPoly(5, 2) = z_draw(0, z0_draw, zmx_draw, scale_draw)
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddPolyline(EnvPoly)
With xxx.Line
    .Weight = 2.5
    .ForeColor.RGB = QBColor(7)
End With
With xxx.Fill
    .ForeColor.RGB = QBColor(7)
    .Transparency = 1
End With
'
' Draw Centerline
'
CtrPoly(1, 1) = x_draw(0, x0_draw, scale_draw)
CtrPoly(1, 2) = z_draw(0 - ZMax * 0.025, z0_draw, zmx_draw, scale_draw)
CtrPoly(2, 1) = x_draw(0, x0_draw, scale_draw)
CtrPoly(2, 2) = z_draw(ZMax * 1.025, z0_draw, zmx_draw, scale_draw)
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddPolyline(CtrPoly)
With xxx.Line
    .Weight = 1#
    .DashStyle = msoLineDashDot
    .ForeColor.RGB = QBColor(1)
End With
'
SLOSH_Sheet.Shapes.SelectAll
Selection.Group.Name = "GroupAll"
SLOSH_Sheet.Shapes.Range(Array("GroupAll")).Select
Selection.ShapeRange.LockAspectRatio = msoTrue
Selection.Copy
Range("G3").Select
ActiveSheet.PasteSpecial Format:="Picture (Microsoft Office Drawing Object)", Link:=False _
   , DisplayAsIcon:=False
Selection.ShapeRange.Name = "GroupPicture"
' Create Chart
Set XChart = SLOSH_Sheet.Shapes.AddChart
XChart.Name = "XChart"
With XChart
    .Width = PicWidth
    .Height = PicHeight
End With
'
ActiveSheet.Shapes.Range(Array("GroupPicture")).Select
XChart.Select
ActiveChart.Paste
'
'Fname = ThisWorkbook.Path & "\SLOSH_GRAPH.jpg"
'Fname = ThisWorkbook.Path & "\SLOSH_GRAPH"
'ActiveChart.Export Filename:=Fname, FilterName:="JPG"
'ActiveChart.Export Filename:=Fname, FilterName:=filtname
Fname = ThisWorkbook.Path & Application.PathSeparator & "SLOSH_GRAPH.bmp"
'ActiveChart.Export Filename:=Fname, FilterName:=filtname
ActiveChart.Export Filename:=Fname
SLOSH_Sheet.Delete
Sheets(SheetName.Text).Activate
'
If Not (ReadSheet) Then Application.ScreenUpdating = True

DrawBox = Fname

End Function



Sub DrawSeg(SegTypeL, RStartL, ZStartL, REndL, ZEndL, ROL, ZOL, RRadiusL, ZRadiusL, np1, Seg1Poly, Fname)
'
'  Draws a tank segment shape
'
Dim SLOSH_Sheet As Worksheet
Dim xxx As Shape
Dim XChart As Shape
Dim np_crv As Integer
Dim dx As Double
Dim zsgn As Integer
Dim IP As Integer
Dim xp As Double
Dim ZP As Double
Dim Arg As Double
Dim i As Integer
Dim filtname As Variant
'
'filtname = "JPG"
filtname = "BMP"

Application.ScreenUpdating = False
Application.DisplayAlerts = False
If ((SheetExists("SLOSH_GRAPH"))) Then
   Sheets("SLOSH_GRAPH").Delete
End If
Set SLOSH_Sheet = Sheets.Add
SLOSH_Sheet.Select
SLOSH_Sheet.Name = "SLOSH_GRAPH"
Sheets("SLOSH_GRAPH").Move after:=Worksheets(Worksheets.Count)
Range("A1").Select

'
'
ReDim pp(1 To np1, 1 To 2) As Single
np1 = 0
np_crv = 100
If SegTypeL = 1 Then
   pp(np1 + 1, 1) = x_draw(RStartL, x0_draw, scale_draw)
   pp(np1 + 1, 2) = z_draw(ZStartL, z0_draw, zmx_draw, scale_draw)
   pp(np1 + 2, 1) = x_draw(REndL, x0_draw, scale_draw)
   pp(np1 + 2, 2) = z_draw(ZEndL, z0_draw, zmx_draw, scale_draw)
   np1 = 2
 Else
   dx = (REndL - RStartL) / (np_crv - 1)
   zsgn = -999
   If ZStartL - ZOL < 0 Then zsgn = -1
   If ZStartL - ZOL > 0 Then zsgn = 1
   If ZStartL - ZOL = 0 Then
      If ZEndL - ZOL < 0 Then zsgn = -1
      If ZEndL - ZOL > 0 Then zsgn = 1
   End If
   For IP = 1 To np_crv
      xp = RStartL + (IP - 1) * dx
      Arg = 1 - ((xp - ROL) / RRadiusL) ^ 2
      If Arg <= 0.00000001 Then Arg = 0
         ZP = ZOL + zsgn * ZRadiusL * Sqr(Arg)
         pp(np1 + IP, 1) = x_draw(xp, x0_draw, scale_draw)
         pp(np1 + IP, 2) = z_draw(ZP, z0_draw, zmx_draw, scale_draw)
   Next IP
   np1 = np_crv
End If
ReDim Seg1Poly(1 To np1, 1 To 2) As Single
For i = 1 To np1
   Seg1Poly(i, 1) = pp(i, 1)
   Seg1Poly(i, 2) = pp(i, 2)
Next i
'
' Make new temporary polyline coordinate set
ReDim pp(1 To NppDraw + np1, 1 To 2)
If NppDraw > 0 Then
   For i = 1 To NppDraw
      pp(i, 1) = SegPoly(i, 1)
      pp(i, 2) = SegPoly(i, 2)
   Next i
End If
For i = 1 To np1
   pp(NppDraw + i, 1) = Seg1Poly(i, 1)
   pp(NppDraw + i, 2) = Seg1Poly(i, 2)
Next i
   
' Redraw the Envelope and the current segments

'
' Add the tank envelope
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddPolyline(EnvPoly)
With xxx.Line
    .Weight = 1.5
    .ForeColor.RGB = QBColor(7)
End With
With xxx.Fill
    .ForeColor.RGB = QBColor(7)
    .Transparency = 1
End With
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddPolyline(CtrPoly)
With xxx.Line
    .Weight = 1#
    .DashStyle = msoLineDashDot
    .ForeColor.RGB = QBColor(1)
End With
'
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddPolyline(pp)
With xxx.Line
    .Weight = 2.5
    .ForeColor.RGB = QBColor(0)
End With
If pp(NppDraw + np1, 1) = pp(1, 1) And pp(NppDraw + np1, 2) = pp(1, 2) Then
   With xxx.Fill
      .ForeColor.RGB = QBColor(7)
      .Transparency = 1
   End With
End If
'
SLOSH_Sheet.Shapes.SelectAll
Selection.Group.Name = "GroupAll"
SLOSH_Sheet.Shapes.Range(Array("GroupAll")).Select
Selection.ShapeRange.LockAspectRatio = msoTrue
Selection.Copy
Range("G3").Select
ActiveSheet.PasteSpecial Format:="Picture (Microsoft Office Drawing Object)", Link:=False _
    , DisplayAsIcon:=False
Selection.ShapeRange.Name = "GroupPicture"
' Create Chart
Set XChart = SLOSH_Sheet.Shapes.AddChart
XChart.Name = "XChart"
With XChart
    .Width = PicWidth
    .Height = PicHeight
End With
'
ActiveSheet.Shapes.Range(Array("GroupPicture")).Select
XChart.Select
ActiveChart.Paste
'
'Fname = ThisWorkbook.Path & "\SLOSH_GRAPH.jpg"
'Fname = ThisWorkbook.Path & "\SLOSH_GRAPH"
'ActiveChart.Export Filename:=Fname, FilterName:="JPG"
Fname = ThisWorkbook.Path & Application.PathSeparator & "SLOSH_GRAPH.bmp"
'ActiveChart.Export Filename:=Fname, FilterName:=filtname
ActiveChart.Export Filename:=Fname
SLOSH_Sheet.Delete
Sheets(SheetName.Text).Activate
'
If Not (ReadSheet) Then Application.ScreenUpdating = True

End Sub



Function DrawLev(ZL_user)
'
'  Draws the liquid level froom the inner radius to the outer radius
'
Dim zl_draw As Double
Dim rlo_draw As Double
Dim rli_draw As Double
Dim SLOSH_Sheet As Worksheet
Dim xxx As Shape
Dim XChart As Shape '
Dim i As Integer
Dim tbox_x As Double
Dim tbox_z As Double
Dim filtname As Variant
'
'filtname = "JPG"
filtname = "BMP"
'
'
' Redraw the tank shape with the liquid line
'
'  Create temporary sheet as workspace
'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
If ((SheetExists("SLOSH_GRAPH"))) Then
   Sheets("SLOSH_GRAPH").Delete
End If
Set SLOSH_Sheet = Sheets.Add
SLOSH_Sheet.Select
SLOSH_Sheet.Name = "SLOSH_GRAPH"
Sheets("SLOSH_GRAPH").Move after:=Worksheets(Worksheets.Count)
Range("A1").Select
'
' Get the liquid surface height
'
zl_draw = ZLiquid
'
'  Find the inner and outer radial locations for the liquid surface
'
rlo_draw = Radius(zl_draw, IOut)
rli_draw = 0
If IIn > 0 Then rli_draw = Radius(zl_draw, IIn)
ReDim pp(1 To 2, 1 To 2) As Single
pp(1, 1) = x_draw(rlo_draw, x0_draw, scale_draw)
pp(1, 2) = z_draw(zl_draw, z0_draw, zmx_draw, scale_draw)
pp(2, 1) = x_draw(rli_draw, x0_draw, scale_draw)
pp(2, 2) = z_draw(zl_draw, z0_draw, zmx_draw, scale_draw)
'
' Redraw the Envelope and the wall segments segments
'
' Tank envelope
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddPolyline(EnvPoly)
With xxx.Line
    .Weight = 1.5
    .ForeColor.RGB = QBColor(7)
End With
With xxx.Fill
    .ForeColor.RGB = QBColor(7)
    .Transparency = 1
End With
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddPolyline(CtrPoly)
With xxx.Line
    .Weight = 1#
    .DashStyle = msoLineDashDot
    .ForeColor.RGB = QBColor(1)
End With
'
' Wall Segments
'
' Make new temporary polyline coordinate set
ReDim ppp(1 To NppDraw, 1 To 2) As Single
For i = 1 To NppDraw
   ppp(i, 1) = SegPoly(i, 1)
   ppp(i, 2) = SegPoly(i, 2)
Next i
'
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddPolyline(ppp)
With xxx.Line
    .Weight = 2.5
    .ForeColor.RGB = QBColor(0)
End With
If ppp(NppDraw, 1) = ppp(1, 1) And ppp(NppDraw, 2) = ppp(1, 2) Then
   With xxx.Fill
      .ForeColor.RGB = QBColor(7)
      .Transparency = 1
   End With
End If
'
'  Add the liquid level
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddPolyline(pp)
With xxx.Line
    .Weight = 1.5
    .ForeColor.RGB = QBColor(3)
End With
'
'  Create a label
tbox_x = x_draw(0, x0_draw, scale_draw)
tbox_z = z_draw(0, z0_draw, zmx_draw, scale_draw) + 20
'
Set xxx = Sheets("SLOSH_GRAPH").Shapes.AddTextbox(msoTextOrientationHorizontal, tbox_x, tbox_z, RMax * scale_draw, 25)
xxx.TextFrame2.TextRange.Characters.Text = "Rough Sketch of Tank"
xxx.TextFrame2.TextRange.Characters(1, 20).ParagraphFormat.FirstLineIndent = 0
With xxx.TextFrame2.TextRange.Characters(1, 20).Font
    .NameComplexScript = "+mn-cs"
    .NameFarEast = "+mn-ea"
    .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
    .Fill.ForeColor.TintAndShade = 0
    .Fill.ForeColor.Brightness = 0
    .Fill.Transparency = 0
    .Fill.Solid
    .Size = 9
    .Name = "+mn-lt"
End With
xxx.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
xxx.TextFrame2.VerticalAnchor = msoAnchorMiddle
xxx.Line.Visible = msoFalse
xxx.Fill.Visible = msoFalse
'
SLOSH_Sheet.Shapes.SelectAll
Selection.Group.Name = "GroupAll"
SLOSH_Sheet.Shapes.Range(Array("GroupAll")).Select
Selection.ShapeRange.LockAspectRatio = msoTrue
Selection.Copy
Range("G3").Select
ActiveSheet.PasteSpecial Format:="Picture (Microsoft Office Drawing Object)", Link:=False _
    , DisplayAsIcon:=False
Selection.ShapeRange.Name = "GroupPicture"
' Create Chart
Set XChart = SLOSH_Sheet.Shapes.AddChart
XChart.Name = "XChart"
With XChart
    .Width = PicWidth
    .Height = PicHeight + 20
End With
'
ActiveSheet.Shapes.Range(Array("GroupPicture")).Select
XChart.Select
ActiveChart.Paste
'
'Fname = ThisWorkbook.Path & "\SLOSH_GRAPH.jpg"
'Fname = ThisWorkbook.Path & "\SLOSH_GRAPH"
'ActiveChart.Export Filename:=Fname, FilterName:="JPG"
Fname = ThisWorkbook.Path & Application.PathSeparator & "SLOSH_GRAPH.bmp"
'ActiveChart.Export Filename:=Fname, FilterName:=filtname
ActiveChart.Export Filename:=Fname
Range("A1").Select
'
SLOSH_Sheet.Delete
'
Sheets(SheetName.Text).Activate
Application.ScreenUpdating = True

End Function



Function FreeSurf()
'
'  Finds the wall segments where the free surface intersects
'  Finds the outer radius of the free surface for
'
Dim ICount As Integer
Dim IFlag As Integer
Dim i As Integer
'
'   --------------------------------------------------------
'   Subroutine determines how many walls (1 or 2) are intersected by the liquid surface and computes the
'   radius of the intersection and which segments are intersected
'   --------------------------------------------------------
'
'       determine how many walls the liquid intersects (1 or 2)
ICount = 0
IFlag = 0
For i = 1 To NSegs
    If ZStart(i) < ZEnd(i) Then
        If ZLiquid > ZStart(i) And ZLiquid <= ZEnd(i) Then
            IFlag = 1
            ICount = ICount + 1
        End If
    Else
        If ZLiquid <= ZStart(i) And ZLiquid > ZEnd(i) Then
            IFlag = 1
            ICount = ICount + 1
        End If
    End If
    If IFlag = 1 Then                           '       liquid intersects this wall so compute radius
        RLiquid(ICount) = Radius(ZLiquid, i)
        IEps(ICount) = i                        '       keep track of which segments are intersected
    End If
    IFlag = 0                                       '       reset the flag
Next i
If ICount = 1 Then                              '       liquid intersects CL and one wall
    Eps = 0                                         '       distance to inner intersection from CL
    RBar = RLiquid(1)                           '       radial length of free surface
'   store segment number of intersected wall
    IOut = IEps(1)
    IIn = 0
Else                                                    '       liquid intersects two walls
    If RLiquid(1) < RLiquid(2) Then
        RBar = RLiquid(2)                       '       radial length of free surface
        Eps = RLiquid(1) / RBar             '       nondimensional distance from CL to inner wall
        '       store segment numbers of the two walls:
        IOut = IEps(2)
        IIn = IEps(1)
    Else
        RBar = RLiquid(1)
        Eps = RLiquid(2) / RBar
        IOut = IEps(1)
        IIn = IEps(2)
    End If
End If

' ZLiquid = ZLiquid / RBar                    '       nondimensionalize

End Function



Function Radius(ZL, N) As Double
'   --------------------------------------------------------------
'   Function determines the radius R coordinate for a given Z coordinate for a segment
'   --------------------------------------------------------------
Dim temp As Double
'
If A3(N) <> 0 Then              '       segment is an arc
    temp = A5(N) + A1(N) * (ZL - ZO(N)) ^ 2
    temp = temp + A2(N) * (ZL - RO(N))
    temp = A4(N) ^ 2 - 4 * A3(N) * temp
    If temp < 0 Then temp = 0 Else temp = Sqr(temp)
        Select Case IQuad(RStart(N), REnd(N), ZStart(N), ZEnd(N), RO(N))
            Case 1, 4
                Radius = RO(N) + (temp - A4(N)) / (2 * A3(N))
            Case 2, 3
                Radius = RO(N) - (temp - A4(N)) / (2 * A3(N))
        End Select
Else                                    '       segment is straight line
    If Abs((A2(N) / A4(N))) < 0.00000001 Then
        Radius = RO(N)
    Else
        Radius = RO(N) - A2(N) * (ZL - ZO(N)) / A4(N)
    End If
End If

End Function



Function IQuad(RS, RE, ZS, ZE, ROC) As Integer
'   --------------------------------------------------------------
'   Function determines for arc segments which quadrant of a circle the arc is in
'   --------------------------------------------------------------
'
If ZE > ZS Then
    If RE > RS Then
        If RE > ROC Then IQuad = 4 Else IQuad = 2
    ElseIf RE < RS Then
            If RE < ROC Then IQuad = 3 Else IQuad = 1
    End If
ElseIf ZE < ZS Then
    If RE > RS Then
        If RE > ROC Then IQuad = 1 Else IQuad = 3
    ElseIf RE < RS Then
        If RE < ROC Then IQuad = 2 Else IQuad = 4
    End If
End If

End Function




Function SheetExists(SheetName As String) As Boolean
'
' Returns TRUE if the sheet exists in the active workbook
'
SheetExists = False
On Error GoTo NoSuchSheet
If Len(Sheets(SheetName).Name) > 0 Then
    SheetExists = True
    Exit Function
End If
NoSuchSheet:
End Function



Function x_draw(xp, x0_draw, scale_draw)
'
' Finds x-position in drawing coordinates
    x_draw = x0_draw + xp * scale_draw
End Function


Function z_draw(ZP, z0_draw, zmx_draw, scale_draw)
'
' Finds z-position in drawing coordinates
    z_draw = z0_draw + zmx_draw - ZP * scale_draw
End Function

Private Sub ReadSheet_Click()

End Sub

Private Sub TankShape_Click()

End Sub
