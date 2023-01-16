Attribute VB_Name = "Elastica"
Option Explicit
Option Base 1
'==============================================================================================================
' Adapted from script from: post by willmac8, #16, on Mar 6, 2014 at:
'   https://www.physicsforums.com/threads/bending-of-a-long-thin-elastic-rod-or-wire-finding-shape-height.735200/
' Also from: https://www.grasshopper3d.com/forum/topics/a-script-for-elastic-bending-aka-the-elastica-curve
'==============================================================================================================
' Original comments...
' -----------------------------------------------------------------
' Elastic Bending Script by Will McElwain
' Created February 2014
'
' DESCRIPTION:
' This beast creates the so-called 'elastica curve', the shape a long, thin rod or wire makes when it is bent elastically (i.e. not permanently). In this case, force
' is assumed to only be applied horizontally (which would be in line with the rod at rest) and both ends are assumed to be pinned or hinged meaning they are free
' to rotate (as opposed to clamped, when the end tangent angle is fixed, usually horizontally). An interesting finding is that it doesn't matter what the material or
' cross-sectional area is, as long as they're uniform along the entire length. Everything makes the same shape when bent as long as it doesn't cross the threshold
' from elastic to plastic (permanent) deformation (I don't bother to find that limit here, but can be found if the yield stress for a material is known).
'
' Key to the formulas used in this script are elliptic integrals, specifically K(m), the complete elliptic integral of the first kind, and E(m), the complete elliptic
' integral of the second kind. There was a lot of confusion over the 'm' and 'k' parameters for these functions, as some people use them interchangeably, but they are
' not the same. m = k^2 (thus k = Sqrt(m)). I try to use the 'm' parameter exclusively to avoid this confusion. Note that there is a unique 'm' parameter for every
' configuration/shape of the elastica curve.
'
' This script tries to find that unique 'm' parameter based on the inputs. The algorithm starts with a test version of m, evaluates an expression, say 2*E(m)/K(m)-1,
' then compares the result to what it should be (in this case, a known width/length ratio). Iterate until the correct m is found. Once we have m, we can then calculate
' all of the other unknowns, then find points that lie on that curve, then interpolate those points for the actual curve. You can also use Wolfram|Alpha as I did to
' find the m parameter based on the equations in this script (example here: http://tiny.cc/t4tpbx for when say width=45.2 and length=67.1).
'
' Other notes:
' * This script works with negative values for width, which will creat a self-intersecting curve (as it should). The curvature of the elastica starts to break down around
' m=0.95 (~154°), but this script will continue to work until M_MAX, m=0.993 (~169°). If you wish to ignore self-intersecting curves, set ignoreSelfIntersecting to True
' * When the only known values are length and height, it is actually possible for certain ratios of height to length to have two valid m values (thus 2 possible widths
' and angles). This script will return them both.
' * Only the first two valid parameters (of the required ones) will be used, meaning if all four are connected (length, width or a PtB, height, and angle), this script will
' only use length and width (or a PtB).
' * Depending on the magnitude of your inputs (say if they're really small, like if length < 10), you might have to increase the constant ROUNDTO at the bottom
'
' REFERENCES:
' {1} "The elastic rod" by M.E. Pacheco Q. & E. Pina, http://www.scielo.org.mx/pdf/rmfe/v53n2/v53n2a8.pdf
' {2} "An experiment in nonlinear beam theory" by A. Valiente, http://www.deepdyve.com/lp/doc/I3lwnxdfGz , also here: http://tiny.cc/Valiente_AEiNBT
' {3} "Snap buckling, writhing and Loop formation In twisted rods" by V.G.A. GOSS, http://myweb.lsbu.ac.uk/~gossga/thesisFinal.pdf
' {4} "Theory of Elastic Stability" by Stephen Timoshenko, http://www.scribd.com/doc/50402462/Timoshenko-Theory-of-Elastic-Stability  (start on p. 76)
'
' INPUT:
' PtA - First anchor point (required)
' PtB - Second anchor point (optional, though 2 out of the 4--length, width, height, angle--need to be specified)
'       [note that PtB can be the same as PtA (meaning width would be zero)]
'       [also note that if a different width is additionally specified that's not equal to the distance between PtA and PtB, then the end point will not equal PtB anymore]
' Pln - Plane of the bent rod/wire, which bends up in the +y direction. The line between PtA and PtB (if specified) must be parallel to the x-axis of this plane
'
' ** 2 of the following 4 need to be specified **
' Len - Length of the rod/wire, which needs to be > 0
' Wid - Width between the endpoints of the curve [note: if PtB is specified in addition, and distance between PtA and PtB <> width, the end point will be relocated
' Ht - Height of the bent rod/wire (when negative, curve will bend downward, relative to the input plane, instead)
' Ang - Inner departure angle or tangent angle (in radians) at the ends of the bent rod/wire. Set up so as width approaches length (thus height approaches zero), angle approaches zero
'
' * Following variables only needed for optional calculating of bending force, not for shape of curve.
' E - Young's modulus (modulus of elasticity) in GPa (=N/m^2) (material-specific. for example, 7075 aluminum is roughly 71.7 GPa)
' I - Second moment of area (or area moment of inertia) in m^4 (cross-section-specific. for example, a hollow rod
'     would have I = pi * (outer_diameter^4 - inner_diameter^4) / 32
' Note: E*I is also known as flexural rigidity or bending stiffness
'
' OUTPUT:
' out - only for debugging messages
' Pts - the list of points that approximate the shape of the elastica
' Crv - the 3rd-degree curve interpolated from those points (with accurate start & end tangents)
' L - the length of the rod/wire
' W - the distance (width) between the endpoints of the rod/wire
' H - the height of the bent rod/wire
' A - the tangent angle at the (start) end of the rod/wire
' F - the force needed to hold the rod/wire in a specific shape (based on the material properties & cross-section) **be sure your units for 'I' match your units for the
' rest of your inputs (length, width, etc.). Also note that the critical buckling load (force) that makes the rod/wire start to bend can be found at height=0
'
' THANKS TO:
' Mårten Nettelbladt (thegeometryofbending.blogspot.com)
' Daniel Piker (Kangaroo plugin)
' David Rutten (Grasshopper guru)
' Euler & Bernoulli (the O.G.'s)
'==============================================================================================================
'
' Note: most of these values for m and h/L ratio were found with Wolfram Alpha and either specific intercepts (x=0) or local minima/maxima. They should be constant.
Public Const M_SKETCHY As Double = 0.95  ' value of the m parameter where the curvature near the ends of the curve gets wonky
Public Const M_MAX As Double = 0.993  ' maximum useful value of the m parameter, above which this algorithm for the form of the curve breaks down
Public Const M_ZERO_W As Double = 0.82611476598497      ' value of the m parameter when width = 0
Public Const M_MAXHEIGHT As Double = 0.701327460663101     ' value of the m parameter at maximum possible height of the bent rod/wire
Public Const M_DOUBLE_W As Double = 0.180254422335014     ' minimum value of the m parameter when two width values are possible for a given height and length
Public Const DOUBLE_W_HL_RATIO As Double = 0.257342117984636     ' value of the height/length ratio above which there are two possible width values
Public Const MAX_HL_RATIO As Double = 0.40314018970565      ' maximum possible value of the height/length ratio

Public Const MAXERR As Double = 0.0000000001  ' error tolerance
Public Const MAXIT As Integer = 100  ' maximum number of iterations
Public Const ROUNDTO As Integer = 10  ' number of decimal places to round off to
Public Const CURVEDIVS As Integer = 50  ' number of sample points for building the curve (or half-curve as it were)

'==============================================================================================================
' For Excel/VBA...
' -----------------------------------------------------------------
' The calculated curve is assumed to have both ends on the x-axis, symetrical, centered on the y-axis
' Values for height (H) and angle (A, radians) must be positive.  Negative width (W) is allowed.
' The four variables L, W, H, and A are used for input and output.  Separate subroutines are provided to
' use any pair of these to calculate the other two parameters.  When the input is L with H, the parameters
' m, W and A may have two sets of values, so global variables mCount, m1, m2, W1, W2, A1, A2 have been added.
' The final "MakeCurve" function from the original script (for producing interpolated points) is not implemented.
' If you need more points, you can increase the value of CURVEDIVS.
'==============================================================================================================

' Added for Excel...
Public Const PI = 3.14159265358979       ' no VBA function for PI (using this instead of worksheet function)
Dim BendFormPointsX As New Collection    ' Create Collection objects for points X and Y coordinates.  Use
Dim BendFormPointsY As New Collection    ' collection instead of array as number of points varies.
Dim mode As String                       ' two characters, defines what is used for input: LW, LH, LA, WH, WA, HA
Dim mCount As Integer
Dim m1 As Double
Dim W1 As Double
Dim A1 As Double
Dim m2 As Double
Dim W2 As Double
Dim A2 As Double

Sub main()
Dim L As Double, W As Double, H As Double, A As Double, F As Double, n As Integer
L = 400  ' cm
W = 318  ' cm (NOTE: units for L, W, H must be consistent)
Call Elastica("LW", L, H, A, W)    ' mode = LW
Call FindBendForm(L, W1, m1, A1)   ' first set of points
For n = 1 To BendFormPointsX.Count
  Debug.Print BendFormPointsX(n), BendFormPointsY(n)
Next
F = Force(130, 0.000000000734, L / 100, m1) ' 9.3 mm diameter Aluminum, Length in meters
Debug.Print L, W1, H, A1, mCount, F

If mCount = 2 Then
  Call FindBendForm(L, W2, m2, A2)   ' second set of points
  For n = 1 To BendFormPointsX.Count
    Debug.Print BendFormPointsX(n), BendFormPointsY(n)
  Next
  Debug.Print L, W2, H, A2
End If
End Sub

Sub MacroLW() ' example for macro to use L and W
' uses specific cells in sheet "Elastica" - see GitHub Readme for layout
Dim L As Double, W As Double, H As Double, A As Double, F As Double, n As Integer
With Worksheets("Elastica")
  'clear results from any previous runs
  .Range("B3").ClearContents
  .Range("B4").ClearContents
  .Range("B9").ClearContents
  .Range(Selection, Selection.End(xlToRight)).Select  'go to end of row
  Selection.ClearContents
  .Range("B8").Select
  .Range(Selection, Selection.End(xlToRight)).Select  'go to end of row
  Selection.ClearContents
  .Range("B7").Select
  'get inputs
  L = .Range("B1") ' cm
  W = .Range("B2") ' cm (NOTE: units for L, W, H must be consistent)
  'run calculation
  Call Elastica("LW", L, H, A, W)    ' mode = LW
  Call FindBendForm(L, W1, m1, A1)   ' generate first set of points
  'outputs
  .Range("B3") = H
  .Range("B4") = A1 * 180 / PI   ' convert output radians to degrees
  For n = 1 To BendFormPointsX.Count
    .Cells(7, 1 + n).Value = BendFormPointsX(n)
    .Cells(8, 1 + n).Value = BendFormPointsY(n)
  Next
  F = Force(.Range("B5"), .Range("B6"), L / 100, m1)
  .Range("B9") = F
  ' ignore second curve result, if any
End With
End Sub

Public Sub Elastica(ByVal mode As String, Optional ByRef L As Variant, Optional ByRef H As Variant, Optional ByVal A As Variant, Optional ByVal W As Variant)
' find parameters m, L, W, H, A for each calculation mode.
' NOTE: m, A and W are returned in global variables m1, m2, A1, A2, W1, W2, with mCount
mCount = 1  ' default

'handle any swapped letters for mode
If mode = "WL" Then mode = "LW"
If mode = "HL" Then mode = "LH"
If mode = "AL" Then mode = "LA"
If mode = "HW" Then mode = "WH"
If mode = "AW" Then mode = "WA"
If mode = "AH" Then mode = "HA"

'calculate for selected mode
Select Case mode
  Case "LW":
    If Not (IsSet(L)) Then
      MsgBox "Length (L) is required for this calculation mode"
    End If
    If Not (IsSet(W)) Then
      MsgBox "Length (W) is required for this calculation mode"
    End If
    If L <= 0 Then
      MsgBox "Length (L) cannot be negative or zero"
      Exit Sub
    End If
    If Abs(W) > L Then
      MsgBox "Width (W) is greater than length (L)"
      Exit Sub
    End If
    If W = L Then ' skip the solver and set the known values
      H = 0
      W1 = W
      A1 = 0
      m1 = 0
    Else
      m1 = SolveMFromLenWid(L, W)   ' only one value for m
      H = Cal_H(L, m1)  ' L * Sqrt(m) / K(m)
      A1 = Cal_A(m1)  ' Acos(1 - 2 * m)
      W1 = W
    End If
  
  Case "LH"
    If Not (IsSet(L)) Then
      MsgBox "Length (L) is required for this calculation mode"
    End If
    If Not (IsSet(H)) Then
      MsgBox "Height (H) is required for this calculation mode"
    End If
    If L <= 0 Then
      MsgBox "Length (L) cannot be negative or zero"
      Exit Sub
    End If
    If Abs(H / L) > MAX_HL_RATIO Then
      MsgBox "Height not possible with given length"
      Exit Sub
    End If
    If H < 0 Then
      H = -H  ' if height is negative, set it to positive
    End If
    If H = 0 Then  ' skip the solver and set the known values
      W1 = L
      A1 = 0
      m1 = 0
    Else
      Call SolveMFromLenHt(L, H)  ' note that it's possible for two values of m to be found if height is close to max height
      ' NOTE: subroutine returns results in global variables: mCount, m1, m2
      If mCount = 1 Then  ' there's only one m value returned
        W1 = Cal_W(L, m1)  ' L * (2 * E(m) / K(m) - 1)
        A1 = Cal_A(m1)  ' Acos(1 - 2 * m)
      End If
      If mCount = 2 Then  ' get second set of W and A values
        W2 = Cal_W(L, m2)  ' L * (2 * E(m) / K(m) - 1)
        A2 = Cal_A(m2)  ' Acos(1 - 2 * m)
      End If
    End If
    
  Case "LA"
    If Not (IsSet(L)) Then
      MsgBox "Length (L) is required for this calculation mode"
    End If
    If Not (IsSet(W)) Then
      MsgBox "Angle (A) is required for this calculation mode"
    End If
    If L <= 0 Then
      MsgBox "Length (L) cannot be negative or zero"
      Exit Sub
    End If
    If A < 0 Then
      A = -A  ' if angle is negative, set it to positive
    End If
    m1 = Cal_M(A)  ' (1 - Cos(a)) / 2
    mCount = 1
    If A = 0 Then  ' skip the solver and set the known values
      W1 = L
      H = 0
    Else
      W1 = Cal_W(L, m1)  ' L * (2 * E(m) / K(m) - 1)
      H = Cal_H(L, m1)  ' L * Sqrt(m) / K(m)
    End If
    A1 = A
    
  Case "WH"
    If Not (IsSet(W)) Then
      MsgBox "Width (W) is required for this calculation mode"
    End If
    If Not (IsSet(H)) Then
      MsgBox "Height (H) is required for this calculation mode"
    End If
    If H < 0 Then
      H = -H  ' if height is negative, set it to positive
    End If
    If H = 0 Then  ' skip the solver and set the known values
      m1 = 0
      L = W
      A1 = 0
    Else
      m1 = SolveMFromWidHt(W, H)
      L = Cal_L(H, m1)  ' h * K(m) / Sqrt(m)
      A1 = Cal_A(m1)  ' Acos(1 - 2 * m)
    End If
    W1 = W
  
  Case "WA"
    If Not (IsSet(W)) Then
      MsgBox "Width (W) is required for this calculation mode"
    End If
    If Not (IsSet(A)) Then
      MsgBox "Angle (A) is required for this calculation mode"
    End If
    If W = 0 Then
      MsgBox "Curve not possible with width = 0 and an angle as inputs"
      Exit Sub
    End If
    If A < 0 Then
      A = -A  ' if angle is negative, set it to positive
    End If
    m1 = Cal_M(A)  ' (1 - Cos(a)) / 2
    If A = 0 Then  ' skip the solver and set the known values
      m1 = 0
      L = W
      H = 0
    Else
      L = W / (2 * EllipticE(m1) / EllipticK(m1) - 1)
      If L < 0 Then
        MsgBox "Curve not possible at specified width and angle (calculated length is negative)"
        Exit Sub
      End If
      H = Cal_H(L, m1)  ' L * Sqrt(m) / K(m)
    End If
    A1 = A
    W1 = W
  
  Case "HA"
    If Not (IsSet(H)) Then
      MsgBox "Height (H) is required for this calculation mode"
    End If
    If Not (IsSet(A)) Then
      MsgBox "Angle (A) is required for this calculation mode"
    End If
    If H < 0 Then
      H = -H  ' if height is negative, set it to positive
    End If
    If H = 0 Then
      MsgBox "Height can't = 0 if only height and angle are specified"
      Exit Sub
    Else
      If A < 0 Then
        A = -A  ' if angle is negative, set it to positive
      End If
      m1 = Cal_M(A)  ' (1 - Cos(a)) / 2
      If A = 0 Then
        MsgBox "Angle can't = 0 if only height and angle are specified"
        Exit Sub
      Else
        L = Cal_L(H, m1)  ' h * K(m) / Sqrt(m)
        W1 = Cal_W(L, m1)  ' L * (2 * E(m) / K(m) - 1)
      End If
    End If
    A1 = A

  Case Else
    MsgBox "Unknown mode: " & mode
    Exit Sub
End Select

End Sub


Public Function Force(ByVal E As Double, ByVal I As Double, ByVal L As Double, ByVal m As Double) As Double
' takes E (Young's Modulus of Elasticity, GPa), I (cross-section area moment of inertia, m^4),
' L (length, meters), and returns Force (Newtons)
' Young's modulus input E is in GPa, so we convert to Pa here (= N/m^2)
Force = (EllipticK(m) ^ 2 * (E * 10 ^ 9) * I / L ^ 2) ' from reference {4} pg. 79
End Function

Sub ClearCollections()
' IMPORTANT! The global collections still hold data after the program ends!
' clear any data accumulated in BendFormPointsX and BendFormPointsY
Dim n As Long
If BendFormPointsX.Count > 0 Then
  For n = 1 To BendFormPointsX.Count
    BendFormPointsX.Remove (1)  ' remove first item
  Next
End If
If BendFormPointsY.Count > 0 Then
  For n = 1 To BendFormPointsY.Count
    BendFormPointsY.Remove (1)  ' remove first item
  Next
End If
End Sub

Private Function IsSet(ByRef param As Variant) As Boolean  ' Check if an input parameter has data
IsSet = False  ' default
If Not (IsMissing(param)) Then
  If IsNumeric(param) Then
    IsSet = True
  End If
End If
End Function

Private Sub msgSub(ByVal msgType As String, ByVal msg As String)  ' Output an error, warning, or informational message
Select Case msgType
  Case "error"
    MsgBox ("Error: " & msg)
    Debug.Print ("Error: " & msg)
  Case "warning"
    MsgBox ("Warning: " & msg)
    Debug.Print ("Warning: " & msg)
  Case "info"
    MsgBox ("Error: " & msg)
    Debug.Print ("Error: " & msg)
End Select
End Sub

' Solve for the m parameter from length and width (reference {1} equation (34), except b = width and K(k) and E(k) should be K(m) and E(m))
Private Function SolveMFromLenWid(ByVal L As Double, ByVal W As Double) As Double
If W = 0 Then
  SolveMFromLenWid = M_ZERO_W  ' for the boundry condition width = 0, bypass the function and return the known m value
End If

Dim n As Integer ' Iteration counter (quit if >MAXIT)
Dim lower As Double ' m must be within this range
Dim upper As Double
Dim m As Double
Dim cwl As Double
n = 1
lower = 0
upper = 1

Do While (((upper - lower) > MAXERR) And (n < MAXIT))  ' Repeat until range narrow enough or MAXIT
  m = (upper + lower) / 2
  cwl = 2 * EllipticE(m) / EllipticK(m) - 1  ' calculate w/L with the test value of m
  If cwl < W / L Then  ' compares the calculated w/L with the actual w/L then narrows the range of possible m
    upper = m
  Else
    lower = m
  End If
  n = 1 + n
Loop
SolveMFromLenWid = m
End Function

Private Sub SolveMFromLenHt(ByVal L As Double, ByVal H As Double)
' Solve for the m parameter from length and height (reference {1} equation (33), except K(k) should be K(m) and k = sqrt(m))
' Note that it's actually possible to find 2 valid values for m (hence 2 width values) at certain height values
' NOTE: subroutine returns results in global variables: mCount, m1, m2

Dim n As Integer ' Iteration counter (quit if >MAXIT)
Dim lower As Double ' m must be within this range
Dim upper As Double
Dim twoWidths As Boolean
Dim m As Double
Dim chl As Double
n = 1
lower = 0
upper = 1
' check to see if h/L is within the range where 2 solutions for the width are possible
If ((H / L >= DOUBLE_W_HL_RATIO) And (H / L < MAX_HL_RATIO)) Then
  twoWidths = True
Else
  twoWidths = False
End If

If twoWidths Then
  ' find the first of two possible solutions for m with the following limits:
  lower = M_DOUBLE_W  ' see constants at bottom of script
  upper = M_MAXHEIGHT  ' see constants at bottom of script
  Do While (((upper - lower) > MAXERR) And (n < MAXIT))  ' Repeat until range narrow enough or MAXIT
    m = (upper + lower) / 2
    chl = Sqr(m) / EllipticK(m)  ' calculate h/L with the test value of m
    If chl > H / L Then  ' compares the calculated h/L with the actual h/L then narrows the range of possible m
      upper = m
    Else
      lower = m
    End If
    n = n + 1
  Loop
  m1 = m
  mCount = 1 ' unless changed below
  ' then look for the second of two possible solutions for m with the following limits:
  lower = M_MAXHEIGHT  ' see constants at bottom of script
  upper = 1
  Do While (((upper - lower) > MAXERR) And (n < MAXIT))  ' Repeat until range narrow enough or MAXIT
    m = (upper + lower) / 2
    chl = Sqr(m) / EllipticK(m)  ' calculate h/L with the test value of m
    If chl < H / L Then  ' compares the calculated h/L with the actual h/L then narrows the range of possible m
      upper = m
    Else
      lower = m
    End If
    n = n + 1
  Loop
  If m <= M_MAX Then  ' return this m parameter only if it falls within the maximum useful value (above which the curve breaks down)
    m2 = m
    mCount = 2
  End If
Else   ' find the one possible solution for the m parameter
  upper = M_DOUBLE_W  ' limit the upper end of the search to the maximum value of m for which only one solution exists
  Do While (((upper - lower) > MAXERR) And (n < MAXIT))  ' Repeat until range narrow enough or MAXIT
    m = (upper + lower) / 2
    chl = Sqr(m) / EllipticK(m)  ' calculate h/L with the test value of m
    If chl > H / L Then  ' compares the calculated h/L with the actual h/L then narrows the range of possible m
      upper = m
    Else
      lower = m
    End If
    n = n + 1
  Loop
  m1 = m
  mCount = 1
End If
End Sub

' Solve for the m parameter from width and height (derived from reference {1} equations (33) and (34) with same notes as above)
Private Function SolveMFromWidHt(ByVal W As Double, ByVal H As Double) As Double
Dim n As Integer ' Iteration counter (quit if >MAXIT)
Dim lower As Double ' m must be within this range
Dim upper As Double
Dim m As Double
Dim cwh As Double
n = 1
lower = 0
upper = 1
Do While (((upper - lower) > MAXERR) And (n < MAXIT))  ' Repeat until range narrow enough or MAXIT
  m = (upper + lower) / 2
  cwh = (2 * EllipticE(m) - EllipticK(m)) / Sqr(m)  ' calculate w/h with the test value of m
  If cwh < W / H Then  ' compares the calculated w/h with the actual w/h then narrows the range of possible m
    upper = m
  Else
    lower = m
  End If
  n = 1 + n
Loop
SolveMFromWidHt = m
End Function

' Calculate length based on height and an m parameter, derived from reference {1} equation (33), except K(k) should be K(m) and k = sqrt(m)
Private Function Cal_L(ByVal H As Double, ByVal m As Double) As Double
  Cal_L = H * EllipticK(m) / Sqr(m)
End Function

' Calculate width based on length and an m parameter, derived from reference {1} equation (34), except b = width and K(k) and E(k) should be K(m) and E(m)
Private Function Cal_W(ByVal L As Double, ByVal m As Double) As Double
  Cal_W = L * (2 * EllipticE(m) / EllipticK(m) - 1)
End Function

' Calculate height based on length and an m parameter, from reference {1} equation (33), except K(k) should be K(m) and k = sqrt(m)
Private Function Cal_H(ByVal L As Double, ByVal m As Double) As Double
  Cal_H = L * Sqr(m) / EllipticK(m)
End Function

' Calculate the unique m parameter based on a start tangent angle, from reference {2}, just above equation (9a), that states k = Sin(angle / 2 + Pi / 4),
' but as m = k^2 and due to this script's need for an angle rotated 90° versus the one in reference {1}, the following formula is the result
' New note: verified by reference {4}, pg. 78 at the bottom
Private Function Cal_M(ByVal A As Double) As Double
  Cal_M = (1 - Cos(A)) / 2   ' equal to Sin^2(a/2) too
End Function

' Calculate start tangent angle based on an m parameter, derived from above formula
Private Function Cal_A(ByVal m As Double) As Double
  Cal_A = Application.WorksheetFunction.Acos(1 - 2 * m)
End Function

' This is the heart of this script, taking the found (or specified) length, width, and angle values along with the found m parameter to create
' a list of points that approximate the shape or form of the elastica. It works by finding the x and y coordinates (which are reversed versus
' the original equations (12a) and (12b) from reference {2} due to the 90° difference in orientation) based on the tangent angle along the curve.
' See reference {2} for more details on how they derived it. Note that to simplify things, the algorithm only calculates the points for half of the
' curve, then mirrors those points along the y-axis.
' the parameter "which" switches betwen two m value curves.  1 for BendFormPointsX and BendFormPointsY, 2 for BendFormPointsXX and BendFormPointsYY

Private Sub FindBendForm(ByVal L As Double, ByVal W As Double, ByVal m As Double, ByVal Ang As Double)
If m >= M_SKETCHY Then
  MsgBox "Accuracy of the curve whose width = " & Round(W, 4) & " is not guaranteed"
End If

L = L / 2  ' because the below algorithm is based on the formulas in reference {2} for only half of the curve
W = W / 2  ' same

Call ClearCollections  ' IMPORTANT! The global collections still hold data after the program ends!
If Ang = 0 Then  ' if angle (and height) = 0, then simply return the start and end points of the straight line
  BendFormPointsX.Add (2 * W)
  BendFormPointsY.Add (0)
  BendFormPointsX.Add (0)
  BendFormPointsY.Add (0)
  Exit Sub
End If

Dim X As Double
Dim Y As Double
Dim halfCurvePtsX As New Collection    ' Create a Collection object for points X coordinate
Dim halfCurvePtsY As New Collection    ' Create a Collection object for points Y coordinate

Ang = Ang - PI / 2  ' a hack to allow this algorithm to work, since the original curve in paper {2} was rotated 90°
Dim angB As Double
angB = Ang + (-PI / 2 - Ang) / CURVEDIVS  ' angB is the 'lowercase theta' which should be in formula {2}(12b) as the interval
' start [a typo...see equation(3)]. It's necessary to start angB at ang + [interval] instead of just ang due to integration failing at angB = ang
halfCurvePtsX.Add (W)   ' start with this known initial point, as integration will fail when angB = ang
halfCurvePtsY.Add (0)

' each point {x, y} is calculated from the tangent angle, angB, that occurs at each point (which is why this iterates from ~ang to -pi/2, the known end condition)
Do While Round(angB, ROUNDTO) >= Round(-PI / 2, ROUNDTO)
  Y = (Sqr(2) * Sqr(Sin(Ang) - Sin(angB)) * (W + L)) / (2 * EllipticE(m))  ' note that x and y are swapped vs. (12a) and (12b)
  X = (L / (Sqr(2) * EllipticK(m))) * Simpson(angB, -PI / 2, 500, Ang)  ' calculate the Simpson approximation of the integral (function f below)
  ' over the interval angB ('lowercase theta') to -pi/2. side note: is 500 too few iterations for the Simson algorithm?
  If Round(X, ROUNDTO) = 0 Then
    X = 0
  End If
  halfCurvePtsX.Add (X)
  halfCurvePtsY.Add (Y)
  angB = angB + (-PI / 2 - Ang) / CURVEDIVS  ' onto the next tangent angle
Loop

' After finding the x and y values for half of the curve, add the {-x, y} values for the rest of the curve
Dim n As Integer, num As Integer
num = halfCurvePtsX.Count
For n = 1 To num
  If Round(halfCurvePtsX(n), ROUNDTO) = 0 Then
    If Round(halfCurvePtsY(n), ROUNDTO) = 0 Then ' special case when width = 0: when x = 0, only duplicate the point when y = 0 too
        BendFormPointsX.Add (0#)
        BendFormPointsY.Add (0#)
    End If
  Else
      BendFormPointsX.Add (-halfCurvePtsX(n))
      BendFormPointsY.Add (halfCurvePtsY(n))
  End If
Next

' add other halfCurvePts in reverse order
For n = num To 1 Step -1
    BendFormPointsX.Add (halfCurvePtsX(n))
    BendFormPointsY.Add (halfCurvePtsY(n))
Next

Set halfCurvePtsX = Nothing  ' may not be required, does not hurt to be explicit
Set halfCurvePtsY = Nothing
End Sub

' Implements the Simpson approximation for an integral of function f below
Public Function Simpson(A As Double, b As Double, n As Integer, theta As Double) As Double 'n should be an even number
Dim j As Integer, s1 As Double, s2 As Double, H As Double
H = (b - A) / n
s1 = 0
s2 = 0
For j = 1 To n - 1 Step 2
  s1 = s1 + fn(A + j * H, theta)
Next j
For j = 2 To n - 2 Step 2
  s2 = s2 + fn(A + j * H, theta)
Next j
Simpson = H / 3 * (fn(A, theta) + 4 * s1 + 2 * s2 + fn(b, theta))
End Function

' Specific calculation for the Simpson approximation integration
Public Function fn(X As Double, theta As Double) As Double
  fn = Sin(X) / (Sqr(Sin(theta) - Sin(X)))  ' from reference {2} formula (12b)
End Function

' Return the Complete Elliptic integral of the 1st kind
' Abramowitz and Stegun p.591, formula 17.3.11
' Code from http://www.codeproject.com/Articles/566614/Elliptic-integrals
Public Function EllipticK(ByVal m As Double) As Double
Dim sum As Double, term As Double, above As Double, below As Double, I As Integer
sum = 1
term = 1
above = 1
below = 2

For I = 1 To 100
  term = term * (above / below)
  sum = sum + (Application.WorksheetFunction.Power(m, I) * Application.WorksheetFunction.Power(term, 2))
  above = 2 + above
  below = 2 + below
Next
sum = sum * 0.5 * PI
EllipticK = sum
End Function

' Return the Complete Elliptic integral of the 2nd kind
' Abramowitz and Stegun p.591, formula 17.3.12
' Code from http://www.codeproject.com/Articles/566614/Elliptic-integrals
Public Function EllipticE(ByVal m As Double) As Double
Dim sum As Double, term As Double, above As Double, below As Double, I As Integer
sum = 1
term = 1
above = 1
below = 2

For I = 1 To 100
  term = term * (above / below)
  sum = sum - (Application.WorksheetFunction.Power(m, I) * Application.WorksheetFunction.Power(term, 2) / above)
  above = 2 + above
  below = 2 + below
Next
sum = sum * 0.5 * PI
EllipticE = sum
End Function
