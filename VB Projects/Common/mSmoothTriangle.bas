Attribute VB_Name = "mSmoothTriangle"
Option Explicit

'mSmoothTriangle 0.8 by dafhi  1/9/2005

'"The New Triangles" smooth-edge triangle
'rasterization framework in visual basic

'This module is in the 'alpha' stage of development.

'BRIEF PROCESS OVERVIEW:

'For instructions, copy/paste RenderExample()
'at the bottom of this module to your form or
'custom render module.

'I break triangles that don't have perfectly horizontal edges into two
'triangles.  One has a flat bottom.  The other has a flat top.

' \`
'  \  `.
'   \____
'    \   `
'     \ `

'The triangle(s) are then broken down into 'left' and 'right' edges,
'fade regions are calculated for each scanline.
'For example: L1 to L2, and R1 to R2.

'I use this comment frequently in code: "For-Next skipover".
'What this means is, I create a condition where once the
'smooth triangle data is calculated, out-of-bounds or other
'"empty" data sets are not calculated because a loop will
'for example encounter something like For X = [50] To [49]
'where loop contents are skipped.

' =============================================

Private Type Tri_OUT
 Wide       As Long
 RectPixels As Long
 Y_LO       As Long
 Y_HI       As Long
 WideM      As Long
 HighM      As Long
 LO_PREV     As Long
 HI_PREV     As Long
End Type
 
Private Type RegionLR
 L1 As Long 'each scanline has two fade regions: Left and Right
 L2 As Long
 R1 As Long
 R2 As Long
End Type

'*******
' PUBLIC
'*******

Public Type FP_NODE
 px As Single
 py As Single
End Type

Public Type ColorMatrix
 NodeUBound As Long
 Pts() As FP_NODE
 ARGB() As Long
End Type

Public TriSurfaceINFO As Tri_OUT

Public triAlpha() As Single 'alpha values between 0 and 1

Public TriBits() As RegionLR
Public TriBitP() As RegionLR

' -- in development --
Public sRed()   As Single
Public sGrn()   As Single
Public sBlu()   As Single
Public sAlp()   As Single
Public LPtrA()  As Long
Public LLPtr    As Long
' --------------------

'***************
' RENDER SURFACE
'***************

Dim SW As Long 'Surface Width
Dim SH As Long 'Surface Height

Dim SH_Top As Long 'Height - 1
Dim SW_Right As Long 'Width - 1

Dim sHmp5 As Single 'SH_Top + 0.5
Dim sWmp5 As Single 'SW_Right + 0.5

'******************
' ALPHA CALCULATION
'******************
Dim area          As Single
Dim dx            As Single
Dim dy            As Single
Dim xi_hi         As Single
Dim xi_lo         As Single
Dim overT         As Single
Dim edgeR         As Single
Dim edgeL         As Single
Dim edgeT         As Single
Dim edgeB         As Single
Dim m_len         As Single

Dim X_ST          As Long
Dim X_ED          As Long
Dim Y_ED          As Long
Dim X1            As Long
Dim X2            As Long
Dim X_M           As Long
Dim X_R           As Long
Dim XA            As Long
Dim YA            As Long
Dim Clip_X1       As Long
Dim Clip_X2       As Long
Dim xTop          As Single 'sorted and clipped endpoints
Dim yTop          As Single '
Dim xBot          As Single '
Dim yBot          As Single '
Dim build_length  As Single

Dim m_IsRight     As Boolean
Dim ySortHigh     As Single
Dim ySortMid      As Single
Dim ySortLow      As Single
Dim xSortHigh     As Single
Dim xSortMid      As Single
Dim xSortLow      As Single
Dim vmid_x_intersect_test As Single
Private Sub m_ClipperShallowNegL(ByVal X1 As Long, ByVal X2 As Long, ByVal Y1 As Long, ByVal abs_sl As Single)
Dim LngX   As Long
Dim ClipL1 As Long
Dim ClipL2 As Long

    abs_sl = dy / -dx
    If Y1 < overT Then
    Else
    End If

    If X1 < X2 Then
        If X1 < X_ED Then
            'x-intercept top of pixel row is to left of either:
            '1. the pixel containing top or left vertex for this edge, or
            '2. the viewport's left pixel column
            X1 = X_ED
        Else
            m_len = X1 + 0.5 - xi_hi
            triAlpha(X1, Y1) = 0.5 * m_len * m_len * abs_sl
            ClipL1 = 1
        End If
        If X2 > X_ST Then
            'x-intercept bottom of pixel row is to right of either:
            '1. the pixel containing bottom or right vertex for this edge, or
            '2. the viewport's right pixel column
            X2 = X_ST
        Else
            m_len = xi_lo - (X2 - 0.5)
            triAlpha(X1, Y1) = 1 - 0.5 * m_len * m_len * abs_sl
            ClipL2 = 1
        End If
        ClipL1 = X1 + ClipL1
        ClipL2 = X2 - ClipL2
        For LngX = ClipL1 To ClipL2
            triAlpha(LngX, Y1) = (LngX - xi_hi) * abs_sl
        Next
    Else 'X1 = X2
        triAlpha(X1, Y1) = X1 + 0.5 - 0.5 * (xi_lo + xi_hi)
    End If 'X1 < X2
    
    TriBits(Y1).L1 = X1
    TriBits(Y1).L2 = X2
    
End Sub
Sub TriSurfaceInitialize(ByVal SurfaceWidth As Integer, ByVal SurfaceHeight As Integer)
Dim Tri1 As Tri_OUT

    'describe draw surface
    SW = SurfaceWidth
    SH = SurfaceHeight
    SW_Right = SW - 1
    SH_Top = SH - 1
    
    sWmp5 = SW - 0.5 'anti-alias screen right
    sHmp5 = SH - 0.5 'top
    
    TriSurfaceINFO = Tri1
    TriSurfaceINFO.Wide = SW
    TriSurfaceINFO.RectPixels = SW * SH
    TriSurfaceINFO.WideM = SW - 1
    TriSurfaceINFO.HighM = SH - 1
    
    If TriSurfaceINFO.RectPixels < 1 Then Exit Sub
    
    Erase triAlpha
    ReDim triAlpha(0 To SW_Right, 0 To SH_Top)
    Erase TriBits
    ReDim TriBits(SH_Top)
    Erase TriBitP
    ReDim TriBitP(SH_Top)
        
End Sub
Sub CalculateRegions(sx1 As Single, sy1 As Single, sx2 As Single, sy2 As Single, sx3 As Single, sy3 As Single)

    TriSurfaceINFO.Y_LO = SH
    TriSurfaceINFO.Y_HI = -1
    
    If sy1 < sy2 Then
        If sy2 < sy3 Then 'V1 is the lowest
            xSortHigh = sx3
            ySortHigh = sy3
            xSortMid = sx2
            ySortMid = sy2
            xSortLow = sx1
            ySortLow = sy1
        Else 'V3 <= V2
            xSortHigh = sx2
            ySortHigh = sy2
            If sy1 < sy3 Then
                xSortMid = sx3
                ySortMid = sy3
                xSortLow = sx1
                ySortLow = sy1
            Else
                xSortMid = sx1
                ySortMid = sy1
                xSortLow = sx3
                ySortLow = sy3
            End If
        End If
    Else 'V2 <= V1
        If sy1 < sy3 Then
            xSortHigh = sx3
            ySortHigh = sy3
            xSortMid = sx1
            ySortMid = sy1
            xSortLow = sx2
            ySortLow = sy2
        Else 'V3 <= V1
            xSortHigh = sx1
            ySortHigh = sy1
            If sy2 < sy3 Then
                xSortMid = sx3
                ySortMid = sy3
                xSortLow = sx2
                ySortLow = sy2
            Else
                xSortMid = sx2
                ySortMid = sy2
                xSortLow = sx3
                ySortLow = sy3
            End If
        End If
    End If

    If ySortLow = ySortHigh Then
        If xSortLow < xSortHigh Then
            dx = xSortLow 'borrowing dx for swap
            xSortLow = xSortHigh
            xSortHigh = dx
        End If
    End If
    If ySortMid = ySortLow Then
        If xSortMid > xSortLow Then
            dx = xSortMid
            xSortMid = xSortLow
            xSortLow = dx
        End If
    End If
    If ySortMid = ySortHigh Then
        If xSortMid < xSortHigh Then
            dx = xSortMid 'borrowing dx for swap
            xSortMid = xSortHigh
            xSortHigh = dx
        End If
    End If
    
    overT = -0.5
    If ySortMid < ySortHigh Then
        vmid_x_intersect_test = xSortLow + (xSortHigh - xSortLow) * (ySortMid - ySortLow) / (ySortHigh - ySortLow)
        If xSortMid < vmid_x_intersect_test Then 'two edges left
            m_FollowEdgeL xSortLow, ySortLow, xSortMid, ySortMid, xSortHigh, ySortHigh
            m_FollowEdgeL xSortMid, ySortMid, xSortHigh, ySortHigh, xSortLow, ySortLow
            overT = -0.5
            m_FollowEdgeR xSortLow, ySortLow, xSortHigh, ySortHigh, xSortMid, ySortMid
        Else 'triangle has no area, or two edges right
            m_FollowEdgeL xSortLow, ySortLow, xSortHigh, ySortHigh, xSortMid, ySortMid
            overT = -0.5
            m_FollowEdgeR xSortLow, ySortLow, xSortMid, ySortMid, xSortHigh, ySortHigh
            m_FollowEdgeR xSortMid, ySortMid, xSortHigh, ySortHigh, xSortLow, ySortLow
        End If 'two edges left
    Else 'flat top
        m_FollowEdgeL xSortLow, ySortLow, xSortHigh, ySortHigh, xSortMid, ySortMid
        overT = -0.5
        m_FollowEdgeR xSortLow, ySortLow, xSortMid, ySortMid, xSortHigh, ySortHigh
        m_FollowEdgeR xSortMid, ySortMid, xSortHigh, ySortHigh, xSortLow, ySortLow
    End If
    
    If TriSurfaceINFO.Y_HI > SH_Top Then TriSurfaceINFO.Y_HI = SH_Top
    
End Sub
Private Sub m_FollowEdgeL(px1_ As Single, py1_ As Single, px2_ As Single, py2_ As Single, ipx_ As Single, ipy_ As Single)
Dim len_h         As Single
Dim len_v         As Single
Dim FoundVertex1_ As Boolean
Dim slope         As Single

    dy = py2_ - py1_ 'should be >= 0
    dx = px2_ - px1_

    If dy <> 0 Then

        m_IsRight = False
        m_Clip_dy_NonZero px1_, py1_, px2_, py2_
        
        If yTop <= yBot Then Exit Sub

        If dx < 0 Then
            
            If yBot < overT Then
                X_M = X1
            Else
                X_M = SW
            End If
            
            slope = dy / dx
            
            m_ClipperNegL slope
            
            If YA < Y_ED Then
            
                For YA = YA + 1 To Y_ED - 1
                    build_length = build_length + 1
                    xi_hi = xBot + dx * build_length / dy
                    X1 = Int(xi_hi + 0.5)
                    X2 = Int(xi_lo + 0.5)
                    If X1 < X2 Then
                        If X1 < X_ED Then
                            X1 = X_ED
                            Clip_X1 = 0
                        Else
                            len_h = X1 + 0.5 - xi_hi
                            triAlpha(X1, YA) = 0.5 * len_h * len_h * -slope
                            Clip_X1 = 1
                        End If
                        If X2 > X_ST Then
                            X2 = X_ST
                            Clip_X2 = 0
                        Else
                            len_h = xi_lo - (X2 - 0.5)
                            triAlpha(X2, YA) = 1 + 0.5 * len_h * len_h * slope
                            Clip_X2 = 1
                        End If
                        For XA = X1 + Clip_X1 To X2 - Clip_X2
                            triAlpha(XA, YA) = (xi_hi - XA) * slope
                        Next
                    Else
                        triAlpha(X1, YA) = X1 + 0.5 - 0.5 * (xi_lo + xi_hi)
                    End If
                    TriBits(YA).L1 = X1
                    TriBits(YA).L2 = X2
                    xi_lo = xi_hi
                Next

                m_ClipperNegL slope
                
            End If 'YA < Y_ED
        
        ElseIf dx = 0 And xBot >= -0.5 And xBot < sWmp5 Or dx <> 0 Then 'dx >= 0
        
            If yBot < overT Then
                X_M = X2
            Else
                X_M = -1
            End If
            
            If dx <> 0 Then slope = dy / dx
            
            m_ClipperNotNegL slope
            
            If YA < Y_ED Then
            
                For YA = YA + 1 To Y_ED - 1
                    build_length = build_length + 1
                    xi_hi = xBot + dx * build_length / dy
                    X2 = Int(xi_hi + 0.5)
                    X1 = Int(xi_lo + 0.5)
                    If X1 < X2 Then
                        If X1 < X_ST Then
                            X1 = X_ST
                            Clip_X1 = 0
                        Else
                            len_h = X1 + 0.5 - xi_lo
                            triAlpha(X1, YA) = 0.5 * len_h * len_h * slope
                            Clip_X1 = 1
                        End If
                        If X2 > X_ED Then
                            X2 = X_ED
                            Clip_X2 = 0
                        Else
                            len_h = xi_hi - (X2 - 0.5)
                            triAlpha(X2, YA) = 1 - 0.5 * len_h * len_h * slope
                            Clip_X2 = 1
                        End If
                        For XA = X1 + Clip_X1 To X2 - Clip_X2
                            triAlpha(XA, YA) = (XA - xi_lo) * slope
                        Next
                    Else
                        triAlpha(X1, YA) = X1 + 0.5 - 0.5 * (xi_lo + xi_hi)
                    End If
                    TriBits(YA).L1 = X1
                    TriBits(YA).L2 = X2
                    xi_lo = xi_hi
                Next
                
                m_ClipperNotNegL slope
                
            End If 'YA < Y_ED
                
        End If 'dx < 0
    
        overT = Y_ED + 0.5
    
    ElseIf dx <> 0 Then

        If py1_ < overT Then
            X_M = X1
        Else
            X_M = -1
        End If
        
        YA = Int(py1_ + 0.5)
        If TriSurfaceINFO.Y_LO > YA Then TriSurfaceINFO.Y_LO = YA
        
        overT = YA + 0.5
        
        If YA > -1 And YA < SH Then
            X1 = LMax(0, Int(px2_ + 0.5))
            X2 = LMin(SW_Right, Int(px1_ + 0.5))
            area = overT - py1_
            len_v = 1 - area
            For XA = X1 To X2
                If XA > X_M Then
                    triAlpha(XA, YA) = area
                Else
                    triAlpha(XA, YA) = triAlpha(XA, YA) - len_v
                End If
            Next
            TriBits(YA).L1 = X1
            TriBits(YA).L2 = LMax(X2, X1 - 1)
        ElseIf YA < 0 Then
            TriSurfaceINFO.Y_LO = 0
        Else
            TriSurfaceINFO.Y_LO = SH
        End If

        overT = YA + 0.5
    
    End If 'dx <> 0 and dy <> 0

    YA = Int(py1_ + 0.5)
    If YA < SH And YA > -1 Then
        XA = Int(px1_ + 0.5)
        If XA < SW And XA > -1 Then
            FoundVertex1_ = True
        End If
    End If
        
    If px1_ = xSortLow And py1_ = ySortLow Then 'And px2_ <> xSortMid And py2_ = ySortMid Then
        If FoundVertex1_ Then
            m_RecompLowOrLowRightVertex px1_, py1_, ipx_, ipy_
        End If
    Else
        If FoundVertex1_ Then
            m_RecompMiddleVertexLeft px1_, py1_, ipx_, ipy_
        End If
    End If
        
End Sub
Private Sub m_ClipperNegL(slope As Single)
Dim len_h As Single

    If dx <> 0 Then slope = dy / dx

    build_length = build_length + 1
    xi_hi = xBot + dx * build_length / dy
    X1 = Int(xi_hi + 0.5)
    X2 = Int(xi_lo + 0.5)
    If X2 < X_ED Then
        X1 = X_ED
    ElseIf X1 > X_ST Then
        X2 = X_ST
    Else
        If X1 < X2 Then
            If X1 < X_ED Then
                X1 = X_ED
                Clip_X1 = 0
            Else
                len_h = X1 + 0.5 - xi_hi
                If X1 < X_M Then
                    triAlpha(X1, YA) = 0.5 * len_h * len_h * -slope
                Else
                    triAlpha(X1, YA) = triAlpha(X1, YA) - (1 + 0.5 * len_h * len_h * slope)
                End If
                Clip_X1 = 1
            End If
            If X2 > X_ST Then
                X2 = X_ST
                Clip_X2 = 0
            Else
                len_h = xi_lo - (X2 - 0.5)
                If X2 < X_M Then
                    triAlpha(X2, YA) = 1 + 0.5 * len_h * len_h * slope
                Else
                    triAlpha(X2, YA) = triAlpha(X2, YA) + 0.5 * len_h * len_h * slope
                End If
                Clip_X2 = 1
            End If
            Clip_X1 = X1 + Clip_X1
            Clip_X2 = X2 - Clip_X2
            For XA = Clip_X1 To Clip_X2
                If XA < X_M Then
                    triAlpha(XA, YA) = (xi_hi - XA) * slope
                Else
                    triAlpha(XA, YA) = triAlpha(XA, YA) - (XA - xi_lo) * slope
                End If
            Next
        Else
            If X1 < X_M Then
                triAlpha(X1, YA) = X1 + 0.5 - 0.5 * (xi_lo + xi_hi)
            Else
                triAlpha(X1, YA) = triAlpha(X1, YA) - (0.5 * (xi_lo + xi_hi) - (X1 - 0.5))
            End If
        End If
    End If
    If X_M = SW Then
        TriBits(YA).L1 = X1
        TriBits(YA).L2 = X2
    Else
        TriBits(YA).L1 = LMin(X1, TriBits(YA).L1)
        TriBits(YA).L2 = LMax(X2, TriBits(YA).L2)
    End If
    xi_lo = xi_hi
    X_M = SW

End Sub
Private Sub m_ClipperNotNegL(slope As Single)
Dim len_h As Single

    If dx <> 0 Then slope = dy / dx

    build_length = build_length + 1
    xi_hi = xBot + dx * build_length / dy
    X2 = Int(xi_hi + 0.5)
    X1 = Int(xi_lo + 0.5)
    If X1 > X_ED Then
        X2 = X_ED
    ElseIf X2 < X_ST Then
        X1 = X_ST
    Else
        If X1 < X2 Then
            If X1 < X_ST Then
                X1 = X_ST
                Clip_X1 = 0
            Else
                len_h = X1 + 0.5 - xi_lo
                If X_M < X1 Then
                    triAlpha(X1, YA) = 0.5 * len_h * len_h * slope
                Else
                    triAlpha(X1, YA) = triAlpha(X1, YA) - (1 - 0.5 * len_h * len_h * slope)
                End If
                Clip_X1 = 1
            End If
            If X2 > X_ED Then
                X2 = X_ED
                Clip_X2 = 0
            Else
                len_h = xi_hi - (X2 - 0.5)
                If X_M < X2 Then
                    triAlpha(X2, YA) = 1 - 0.5 * len_h * len_h * slope
                Else
                    triAlpha(X2, YA) = triAlpha(X2, YA) - 0.5 * len_h * len_h * slope
                End If
                Clip_X2 = 1
            End If
            Clip_X1 = X1 + Clip_X1
            Clip_X2 = X2 - Clip_X2
            For XA = Clip_X1 To Clip_X2
                If X_M < XA Then
                    triAlpha(XA, YA) = (XA - xi_lo) * slope
                Else
                    triAlpha(XA, YA) = triAlpha(XA, YA) - (xi_hi - XA) * slope
                End If
            Next
        Else
            If X_M < X1 Then
                triAlpha(X1, YA) = X1 + 0.5 - 0.5 * (xi_lo + xi_hi)
            Else
                triAlpha(X1, YA) = triAlpha(X1, YA) - (0.5 * (xi_lo + xi_hi) - (X1 - 0.5))
            End If
        End If
    End If
    If X_M > -1 Then
        TriBits(YA).L1 = LMin(X1, TriBits(YA).L1)
        TriBits(YA).L2 = LMax(X2, TriBits(YA).L2)
    Else
        TriBits(YA).L1 = X1
        TriBits(YA).L2 = X2
    End If
    xi_lo = xi_hi
    X_M = -1

End Sub
Private Sub m_FollowEdgeR(px1_ As Single, py1_ As Single, px2_ As Single, py2_ As Single, ipx_ As Single, ipy_ As Single)
Dim L1            As Long
Dim L2            As Long
Dim R1            As Long
Dim R2            As Long
Dim len_h         As Single
Dim len_v         As Single
Dim FoundVertex2_ As Boolean
Dim slope         As Single

    dy = py2_ - py1_ 'should be >= 0
    dx = px2_ - px1_
        
    If dy <> 0 Then
        
        m_IsRight = True
        m_Clip_dy_NonZero px1_, py1_, px2_, py2_
        
        If yTop <= yBot Then Exit Sub

        If dx < 0 Then
            If yBot < overT Then
                R1 = TriBits(YA).R1
            Else
                R1 = SW
            End If
            slope = dy / dx
            m_ClipperNegR L2, R1, slope
            If YA < Y_ED Then
                For YA = YA + 1 To Y_ED - 1
                    build_length = build_length + 1
                    xi_hi = xBot + dx * build_length / dy
                    X1 = Int(xi_hi + 0.5)
                    X2 = Int(xi_lo + 0.5)
                    L2 = TriBits(YA).L2
                    If X1 < X2 Then
                        If X1 < X_ED Then
                            X1 = X_ED
                            Clip_X1 = 0
                        Else
                            len_h = X1 + 0.5 - xi_hi
                            If X1 <= L2 Or X1 >= R1 Then
                                triAlpha(X1, YA) = triAlpha(X1, YA) + 0.5 * len_h * len_h * slope
                            Else
                                triAlpha(X1, YA) = 1 + 0.5 * len_h * len_h * slope
                            End If
                            Clip_X1 = 1
                        End If
                        If X2 > X_ST Then
                            X2 = X_ST
                            Clip_X2 = 0
                        Else
                            len_h = xi_lo - (X2 - 0.5)
                            If X2 <= L2 Or X2 >= R1 Then
                                triAlpha(X2, YA) = triAlpha(X2, YA) - (1 + 0.5 * len_h * len_h * slope)
                            Else
                                triAlpha(X2, YA) = 0.5 * len_h * len_h * dy / -dx
                            End If
                            Clip_X2 = 1
                        End If
                        For XA = X1 + Clip_X1 To X2 - Clip_X2
                            If XA <= L2 Or XA >= R1 Then
                                triAlpha(XA, YA) = triAlpha(XA, YA) - (xi_hi - XA) * slope
                            Else
                                triAlpha(XA, YA) = (XA - xi_lo) * slope
                            End If
                        Next
                    Else
                        If X1 <= L2 Or X1 >= R1 Then
                            triAlpha(X1, YA) = triAlpha(X1, YA) - (X1 + 0.5 - 0.5 * (xi_lo + xi_hi))
                        Else
                            triAlpha(X1, YA) = 0.5 * (xi_lo + xi_hi) - (X1 - 0.5)
                        End If
                    End If
                    If X1 <= L2 Then X1 = L2 + 1
                    TriBits(YA).R1 = X1
                    TriBits(YA).R2 = X2
                    xi_lo = xi_hi
                Next
                m_ClipperNegR L2, R1, slope
            End If
        ElseIf dx = 0 And xBot >= -0.5 And xBot < sWmp5 Or dx <> 0 Then 'dx >= 0
            If yBot < overT Then
                X_R = LMax(X2, TriBits(YA).R2)
            Else
                X_R = -1
            End If
            If dx <> 0 Then slope = dy / dx
            m_ClipperNotNegR L2, R1, slope
            If YA < Y_ED Then
                For YA = YA + 1 To Y_ED - 1
                    build_length = build_length + 1
                    xi_hi = xBot + dx * build_length / dy
                    X2 = Int(xi_hi + 0.5)
                    X1 = Int(xi_lo + 0.5)
                    L2 = TriBits(YA).L2
                    X_M = LMax(X_R, L2)
                    If X1 < X2 Then
                        If X1 < X_ST Then
                            X1 = X_ST
                            Clip_X1 = 0
                        Else
                            len_h = X1 + 0.5 - xi_lo
                            If X1 <= X_M Then
                                triAlpha(X1, YA) = triAlpha(X1, YA) - 0.5 * len_h * len_h * slope
                            Else
                                triAlpha(X1, YA) = 1 - 0.5 * len_h * len_h * slope
                            End If
                            Clip_X1 = 1
                        End If
                        If X2 > X_ED Then
                            X2 = X_ED
                            Clip_X2 = 0
                        Else
                            len_h = xi_hi - (X2 - 0.5)
                            If X2 <= X_M Then
                                triAlpha(X2, YA) = triAlpha(X2, YA) - (1 - 0.5 * len_h * len_h * slope)
                            Else
                                triAlpha(X2, YA) = 0.5 * len_h * len_h * slope
                            End If
                            Clip_X2 = 1
                        End If
                        For XA = X1 + Clip_X1 To X2 - Clip_X2
                            If XA > X_M Then
                                triAlpha(XA, YA) = (xi_hi - XA) * slope
                            Else
                                triAlpha(XA, YA) = triAlpha(XA, YA) - (XA - xi_lo) * slope
                            End If
                        Next
                    Else
                        If X1 > X_M Then
                            triAlpha(X1, YA) = 0.5 * (xi_lo + xi_hi) - (X1 - 0.5)
                        Else
                            triAlpha(X1, YA) = triAlpha(X1, YA) - (X1 + 0.5 - 0.5 * (xi_lo + xi_hi))
                        End If
                    End If
                    If X1 <= L2 Then X1 = L2 + 1
                    TriBits(YA).R1 = X1
                    TriBits(YA).R2 = X2
                    xi_lo = xi_hi
                Next
                m_ClipperNotNegR L2, R1, slope
            End If
        End If 'dx < 0
    
        overT = Y_ED + 0.5
        
    ElseIf dx <> 0 Then
        
        YA = Int(py1_ + 0.5)
        If TriSurfaceINFO.Y_HI < YA Then TriSurfaceINFO.Y_HI = YA
        
        If YA > -1 And YA < SH Then
            L2 = TriBits(YA).L2
            R1 = TriBits(YA).R1
            X1 = LMax(0, Int(px2_ + 0.5))
            X2 = LMin(SW_Right, Int(px1_ + 0.5))
            area = py1_ - (YA - 0.5)
            len_v = 1 - area
            For XA = X1 To X2
                If XA <= L2 Or XA >= R1 Then
                    triAlpha(XA, YA) = triAlpha(XA, YA) - len_v
                Else
                    triAlpha(XA, YA) = area
                End If
            Next
            TriBits(YA).R1 = LMax(X1, L2 + 1)
            TriBits(YA).R2 = LMax(X2, TriBits(YA).R2)
        ElseIf YA < 0 Then
            TriSurfaceINFO.Y_HI = -1
        Else
            TriSurfaceINFO.Y_HI = SH_Top
        End If
        
        overT = YA + 0.5
        
    End If 'dy = 0
        
    YA = Int(py2_ + 0.5)
    If YA > -1 And YA < SH Then
        XA = Int(px2_ + 0.5)
        If XA > -1 And XA < SW Then
            FoundVertex2_ = True
        End If
    End If
        
    If px2_ = xSortMid And py2_ = ySortMid Then
        If FoundVertex2_ Then
            m_RecompMiddleVertexRight px2_, py2_, ipx_, ipy_
        End If
    ElseIf px2_ = xSortHigh And py2_ = ySortHigh Then
        If FoundVertex2_ Then
            m_RecompTopOrTopleftVertex px2_, py2_, ipx_, ipy_
        End If
    End If

End Sub
Private Sub m_ClipperNegR(L2 As Long, R1 As Long, slope As Single)
Dim len_h As Single

    build_length = build_length + 1
    xi_hi = xBot + dx * build_length / dy
    X1 = Int(xi_hi + 0.5)
    X2 = Int(xi_lo + 0.5)
    L2 = TriBits(YA).L2
    If X2 < X_ED Then
        X1 = X_ED
    ElseIf X1 > X_ST Then
        X2 = X_ST
    Else
        If X1 < X2 Then
            If X1 < X_ED Then
                X1 = X_ED
                Clip_X1 = 0
            Else
                len_h = X1 + 0.5 - xi_hi
                If X1 <= L2 Or X1 >= R1 Then
                    triAlpha(X1, YA) = triAlpha(X1, YA) + 0.5 * len_h * len_h * dy / dx
                Else
                    triAlpha(X1, YA) = 1 + 0.5 * len_h * len_h * dy / dx
                End If
                Clip_X1 = 1
            End If
            If X2 > X_ST Then
                X2 = X_ST
                Clip_X2 = 0
            Else
                len_h = xi_lo - (X2 - 0.5)
                If X2 <= L2 Or X2 >= R1 Then
                    triAlpha(X2, YA) = triAlpha(X2, YA) - (1 + 0.5 * len_h * len_h * dy / dx)
                Else
                    triAlpha(X2, YA) = 0.5 * len_h * len_h * dy / -dx
                End If
                Clip_X2 = 1
            End If
            For XA = X1 + Clip_X1 To X2 - Clip_X2
                If XA <= L2 Or XA >= R1 Then
                    triAlpha(XA, YA) = triAlpha(XA, YA) - (xi_hi - XA) * dy / dx '?
                Else
                    triAlpha(XA, YA) = (XA - xi_lo) * dy / dx
                End If
            Next
        Else
            If X1 <= L2 Or X1 >= R1 Then
                triAlpha(X1, YA) = triAlpha(X1, YA) - (X1 + 0.5 - 0.5 * (xi_lo + xi_hi))
            Else
                triAlpha(X1, YA) = 0.5 * (xi_lo + xi_hi) - (X1 - 0.5)
            End If
        End If
    End If
    If X1 <= L2 Then X1 = L2 + 1
    If R1 = SW Then
        TriBits(YA).R1 = X1
        TriBits(YA).R2 = X2
    Else
        TriBits(YA).R1 = LMin(X1, TriBits(YA).R1)
        TriBits(YA).R2 = LMax(X2, TriBits(YA).R2)
    End If
    xi_lo = xi_hi
    R1 = SW

End Sub
Private Sub m_ClipperNotNegR(L2 As Long, R1 As Long, slope As Single)
Dim len_h As Single

    build_length = build_length + 1
    xi_hi = xBot + dx * build_length / dy
    X2 = Int(xi_hi + 0.5)
    X1 = Int(xi_lo + 0.5)
    L2 = TriBits(YA).L2
    If X1 > X_ED Then
        X2 = X_ED
    ElseIf X2 < X_ST Then
        X1 = X_ST
    Else
        X_M = LMax(X_R, L2)
        If X1 < X2 Then
            If X1 < X_ST Then
                X1 = X_ST
                Clip_X1 = 0
            Else
                len_h = X1 + 0.5 - xi_lo
                If X1 <= X_M Then
                    triAlpha(X1, YA) = triAlpha(X1, YA) - 0.5 * len_h * len_h * slope
                Else
                    triAlpha(X1, YA) = 1 - 0.5 * len_h * len_h * slope
                End If
                Clip_X1 = 1
            End If
            If X2 > X_ED Then
                X2 = X_ED
                Clip_X2 = 0
            Else
                len_h = xi_hi - (X2 - 0.5)
                If X2 <= X_M Then
                    triAlpha(X2, YA) = triAlpha(X2, YA) - (1 - 0.5 * len_h * len_h * slope)
                Else
                    triAlpha(X2, YA) = 0.5 * len_h * len_h * slope
                End If
                Clip_X2 = 1
            End If
            For XA = X1 + Clip_X1 To X2 - Clip_X2
                If XA > X_M Then
                    triAlpha(XA, YA) = (xi_hi - XA) * slope
                Else
                    triAlpha(XA, YA) = triAlpha(XA, YA) - (XA - xi_lo) * slope
                End If
            Next
        Else
            If X1 > X_M Then
                triAlpha(X1, YA) = 0.5 * (xi_lo + xi_hi) - (X1 - 0.5)
            Else
                triAlpha(X1, YA) = triAlpha(X1, YA) - (X1 + 0.5 - 0.5 * (xi_lo + xi_hi))
            End If
        End If
    End If
    If X1 <= L2 Then X1 = L2 + 1
    If X_R = -1 Then
        TriBits(YA).R1 = X1
        TriBits(YA).R2 = X2
    Else
        TriBits(YA).R1 = LMin(X1, TriBits(YA).R1)
        TriBits(YA).R2 = X2
    End If
    xi_lo = xi_hi
    X_R = -1

End Sub
Private Sub m_RecompLowOrLowRightVertex(vertex_x!, vertex_y!, ipx_!, ipy_!)
Dim dx_    As Single
Dim dy_    As Single
Dim xi_l1 As Single
Dim xi_lo_ As Single
Dim area1  As Single
Dim area2  As Single
Dim len_h  As Single
Dim len_v  As Single
Dim len_1  As Single

    edgeR = XA + 0.5
    edgeL = edgeR - 1
    edgeB = YA - 0.5
    
    len_h = edgeR - vertex_x
    len_v = vertex_y - edgeB
    
    'Now compute the overlap (based on triangle's left edge)
    
    dx_ = ipx_ - vertex_x
    dy_ = ipy_ - vertex_y
    
    If dy_ <> 0 Or dx_ <> 0 Then
    
        If dy <> 0 And dx <> 0 Then
            xi_l1 = vertex_x - dx * len_v / dy
            If xi_l1 >= edgeR Then 'SCENARIO 1
                '|           |
                '|           |
                '|           |
                '|      -----| - -
                '|        ---|     len_1
                '|          -| - -
                '|           |
                '|           |
                len_1 = len_h * dy / -dx
                area1 = 0.5 * len_h * len_1
            ElseIf xi_l1 < edgeL Then 'SCENARIO 2
                '|       |   |
                '| len_1 |   |
                '|       |   |
                '|<----->+---| ----
                '|     ------|  ^
                '|  ---------|  | len_v = rectangular area
                '|-----------|  v
                '|-----------| ----
                len_1 = vertex_x - edgeL
                area1 = len_v - 0.5 * len_1 * len_1 * dy / dx 'rectangle - triangle
            Else 'SCENARIO 3
                area1 = len_v * (edgeR - 0.5 * (xi_l1 + vertex_x)) 'trapezoid
            End If
        ElseIf dy = 0 And dx = 0 Then
            'nothing
        ElseIf dx = 0 Then
            area1 = len_v * len_h
        End If
        
        If dx_ = 0 Then
            area2 = len_v * len_h
        ElseIf dy_ <> 0 Then
            xi_lo_ = vertex_x - dx_ * len_v / dy_
            If xi_lo_ >= edgeR Then 'SCENARIO 1
                area2 = 0.5 * len_h * len_h * dy_ / -dx_
            ElseIf xi_lo_ < edgeL Then 'SCENARIO 2
                len_1 = vertex_x - edgeL
                area2 = len_v - 0.5 * len_1 * len_1 * dy_ / dx_
            Else 'SCENARIO 3
                area2 = len_v * (edgeR - 0.5 * (xi_lo_ + vertex_x))
            End If
        End If
    End If
    
    triAlpha(XA, YA) = triAlpha(XA, YA) + area2 - area1
    
End Sub




Private Sub m_RecompTopOrTopleftVertex(vertex_x!, vertex_y!, ipx_!, ipy_!)
Dim dx_    As Single
Dim dy_    As Single
Dim xi_hi_ As Single
Dim xi_h1  As Single
Dim area1  As Single
Dim area2  As Single
Dim len_h  As Single
Dim len_v  As Single
Dim len_hR As Single

    edgeT = YA + 0.5
    edgeL = XA - 0.5
    edgeR = edgeL + 1
    
    len_v = edgeT - vertex_y
    len_h = vertex_x - edgeL
    len_hR = edgeR - vertex_x
    
    If dy <> 0 And dx <> 0 Then
        xi_h1 = vertex_x + dx * len_v / dy
        If xi_h1 >= edgeR Then
            area1 = 0.5 * len_hR * len_hR * dy / dx
        ElseIf xi_h1 < edgeL Then
            area1 = len_v + 0.5 * len_h * len_h * dy / dx
        Else
            area1 = len_v * (edgeR - 0.5 * (vertex_x + xi_h1))
        End If
    ElseIf dx = 0 And dy = 0 Then
        'nothing
    ElseIf dy = 0 Then
        area1 = len_v
    Else
        area1 = len_v * len_hR
    End If
    
    dy_ = vertex_y - ipy_
    dx_ = vertex_x - ipx_
    
    If dx_ <> 0 And dy_ <> 0 Then
        xi_hi_ = vertex_x + dx_ * len_v / dy_
        If xi_hi_ >= edgeR Then
            area2 = 0.5 * len_hR * len_hR * dy_ / dx_
        ElseIf xi_hi_ < edgeL Then
            area2 = len_v + 0.5 * len_h * len_h * dy_ / dx_
        Else
            area2 = len_v * (edgeR - 0.5 * (vertex_x + xi_hi_))
        End If
    ElseIf dx_ = 0 And dy_ = 0 Then
        'nothing
    ElseIf dx_ = 0 Then
        area2 = len_v * len_hR
    Else
        area2 = len_v
    End If
    
    triAlpha(XA, YA) = triAlpha(XA, YA) + area1 - area2
    
End Sub

Private Sub m_RecompMiddleVertexRight(vertex_x!, vertex_y!, ipx_!, ipy_!)
Dim dx_    As Single
Dim dy_    As Single
Dim xi_lo_ As Single
Dim xi_h1  As Single
Dim area1  As Single
Dim area2  As Single
Dim len_h  As Single
Dim len_v  As Single
Dim len_hL As Single

    edgeR = XA + 0.5
    edgeL = edgeR - 1
    edgeT = YA + 0.5
    
    len_h = edgeR - vertex_x
    len_v = edgeT - vertex_y
    
    len_hL = vertex_x - edgeL
    
    dx_ = ipx_ - vertex_x
    dy_ = ipy_ - vertex_y

    If dx_ <> 0 Or dy_ <> 0 Then
    
        If dx <> 0 And dy <> 0 Then
            xi_h1 = vertex_x + dx * len_v / dy
            If xi_h1 >= edgeR Then
                area1 = 0.5 * len_h * len_h * dy / dx
            ElseIf xi_h1 < edgeL Then
                area1 = len_v + 0.5 * len_hL * len_hL * dy / dx
            Else
                area1 = len_v * (edgeR - 0.5 * (vertex_x + xi_h1))
            End If
        ElseIf dx = 0 And dy = 0 Then
            'nothing
        ElseIf dx = 0 Then
            area1 = len_h * len_v
        Else
            area1 = len_v
        End If
        
        If dx_ = 0 Then
            area2 = len_h * (1 - len_v)
        ElseIf dy_ <> 0 Then
            If dx <> 0 Or dy <> 0 Then
                xi_lo_ = vertex_x - dx_ * (1 - len_v) / dy_
                If xi_lo_ >= edgeR Then
                    area2 = 0.5 * len_h * len_h * dy_ / -dx_
                ElseIf xi_lo_ < edgeL Then
                    area2 = (1 - len_v) - 0.5 * len_hL * len_hL * dy_ / dx_
                Else
                    area2 = (1 - len_v) * (edgeR - 0.5 * (vertex_x + xi_lo_))
                End If
            End If
        End If
        
    End If
    
    area = triAlpha(XA, YA) + area1 + area2
    
    triAlpha(XA, YA) = area
    
End Sub
Private Sub m_RecompMiddleVertexLeft(vertex_x!, vertex_y!, ipx_!, ipy_!)
Dim dx_    As Single
Dim dy_    As Single
Dim xi_hi_ As Single
Dim xi_l1  As Single
Dim area1  As Single
Dim area2  As Single
Dim len_h  As Single
Dim len_v  As Single
Dim len_hL As Single

    edgeR = XA + 0.5
    edgeL = edgeR - 1
    edgeT = YA + 0.5
    
    len_h = edgeR - vertex_x
    len_v = edgeT - vertex_y
    
    len_hL = vertex_x - edgeL
    
    If dx <> 0 And dy <> 0 Then
         xi_l1 = vertex_x - dx * (1 - len_v) / dy
        If xi_l1 >= edgeR Then
            area1 = 1 - len_v + 0.5 * len_h * len_h * dy / dx 'dx is negative
        ElseIf xi_l1 < edgeL Then
            area1 = 0.5 * len_hL * len_hL * dy / dx
        Else
            area1 = (1 - len_v) * (0.5 * (vertex_x + xi_l1) - edgeL)
        End If
    ElseIf dx = 0 And dy = 0 Then
        'nothing
    ElseIf dx = 0 Then
        area1 = (1 - len_v) * len_hL
    Else
        'nothing
    End If

    dx_ = vertex_x - ipx_
    dy_ = vertex_y - ipy_
    
    If dy_ <> 0 And dx_ <> 0 Then
    
        xi_hi_ = vertex_x + dx_ * len_v / dy_
        If xi_hi_ >= edgeR Then
            area2 = len_v - 0.5 * len_h * len_h * dy_ / dx_
        ElseIf xi_hi_ < edgeL Then
            area2 = 0.5 * len_hL * len_hL * dy_ / -dx_
        Else
            area2 = len_v * (0.5 * (vertex_x + xi_hi_) - edgeL)
        End If
    ElseIf dx_ = 0 And dy_ = 0 Then
        'nothing
    ElseIf dx_ = 0 Then
        area2 = len_v * len_hL
    Else
        'nothing
    End If
    
    area = triAlpha(XA, YA) + area1 + area2
    
    triAlpha(XA, YA) = area

End Sub


Private Sub m_Clip_dy_NonZero(pxBot_!, pyBot_!, pxTop_!, pyTop_!) ', BotNotClp As Boolean, TopNotClp As Boolean)

    YA = Int(pyBot_ + 0.5)

    If dx > 0 Then

        If pyTop_ >= sHmp5 Then 'vertex >= screen top
            xTop = pxTop_ - dx * (pyTop_ - sHmp5) / dy 'find where edge hits top of screen
            If xTop >= sWmp5 Then 'intercept is to the right of screen edge
                xTop = sWmp5
                If dx <> 0 Then
                    yTop = pyTop_ + dy * (xTop - pxTop_) / dx
                Else
                    yTop = sHmp5
                End If
                If Not m_IsRight Then TriSurfaceINFO.Y_HI = Int(yTop + 0.5)
                m_RegionReset LMax(yTop, overT), pyTop_, True
            Else
                yTop = sHmp5
                If Not m_IsRight Then TriSurfaceINFO.Y_HI = SH_Top
            End If
        ElseIf pxTop_ >= sWmp5 Then
            xTop = sWmp5
            yTop = pyTop_
            If dx <> 0 Then yTop = yTop + dy * (xTop - pxTop_) / dx
            m_RegionReset LMax(yTop, overT), pyTop_, True
            If Not m_IsRight Then TriSurfaceINFO.Y_HI = Int(yTop + 0.5)
        Else
            xTop = pxTop_
            yTop = pyTop_
            If Not m_IsRight Then If Int(yTop + 0.5) > TriSurfaceINFO.Y_HI Then TriSurfaceINFO.Y_HI = Int(yTop + 0.5)
        End If
    
        If pyBot_ < -0.5 Then
            xBot = pxBot_ + dx * (-0.5 - pyBot_) / dy
            If xBot < -0.5 Then
                xBot = -0.5
                If dx <> 0 Then
                    yBot = pyBot_ - dy * (pxBot_ - xBot) / dx
                Else
                    yBot = -0.5
                End If
                m_RegionReset overT, yBot '- 1
                If m_IsRight Then TriSurfaceINFO.Y_LO = Int(yBot + 0.5)
                XA = -1
            Else
                yBot = -0.5
                If m_IsRight Then TriSurfaceINFO.Y_LO = 0
                XA = Int(xBot + 0.5)
            End If
        ElseIf pxBot_ < -0.5 Then
            xBot = -0.5
            yBot = pyBot_
            If dx <> 0 Then yBot = yBot + dy * (xBot - pxBot_) / dx
            m_RegionReset LMax(overT, pyBot_), yBot '- 1
            If m_IsRight Then TriSurfaceINFO.Y_LO = Int(yBot + 0.5)
            XA = -1
        Else
            xBot = pxBot_
            yBot = pyBot_
            If Not m_IsRight Then If Int(yBot + 0.5) < TriSurfaceINFO.Y_LO Then TriSurfaceINFO.Y_LO = Int(yBot + 0.5)
            XA = Int(xBot + 0.5)
        End If
    
    Else 'dx <= 0
    
        If pyTop_ >= sHmp5 Then 'vertex >= screen top
            xTop = pxTop_ - dx * (pyTop_ - sHmp5) / dy 'find where edge hits top of screen
            If xTop < -0.5 Then 'intercept is to the left of screen edge
                xTop = -0.5
                If dx <> 0 Then
                    yTop = pyTop_ + dy * (xTop - pxTop_) / dx
                Else
                    yTop = sHmp5
                End If
                m_RegionReset LMax(yTop, overT), pyTop_
                If m_IsRight Then TriSurfaceINFO.Y_HI = Int(yTop + 0.5)
            Else
                yTop = sHmp5
                If m_IsRight Then TriSurfaceINFO.Y_HI = SH_Top
            End If
        ElseIf pxTop_ < -0.5 Then
            xTop = -0.5
            yTop = pyTop_
            If dx <> 0 Then yTop = yTop + dy * (xTop - pxTop_) / dx
            m_RegionReset LMax(yTop, overT), pyTop_
            If m_IsRight Then TriSurfaceINFO.Y_HI = Int(yTop + 0.5)
        Else
            xTop = pxTop_
            yTop = pyTop_
            If Not m_IsRight Then If Int(yTop + 0.5) > TriSurfaceINFO.Y_HI Then TriSurfaceINFO.Y_HI = Int(yTop + 0.5)
        End If
    
        If pyBot_ < -0.5 Then
            xBot = pxBot_ + dx * (-0.5 - pyBot_) / dy
            If xBot >= sWmp5 Then
                xBot = sWmp5
                If dx <> 0 Then
                    yBot = pyBot_ + dy * (xBot - pxBot_) / dx
                Else
                    yBot = -0.5
                End If
                m_RegionReset overT, yBot, True
                If Not m_IsRight Then TriSurfaceINFO.Y_LO = Int(yBot + 0.5)
                XA = SW
            Else
                yBot = -0.5
                If Not m_IsRight Then TriSurfaceINFO.Y_LO = 0
                XA = Int(xBot + 0.5)
            End If
        ElseIf pxBot_ >= sWmp5 Then
            xBot = sWmp5
            yBot = pyBot_
            If dx <> 0 Then yBot = yBot - dy * (pxBot_ - xBot) / dx
            m_RegionReset overT, yBot, True
            If Not m_IsRight Then If Int(yBot + 0.5) < TriSurfaceINFO.Y_LO Then TriSurfaceINFO.Y_LO = Int(yBot + 0.5)
            XA = SW
        Else
            xBot = pxBot_
            yBot = pyBot_
            If Not m_IsRight Then If Int(yBot + 0.5) < TriSurfaceINFO.Y_LO Then TriSurfaceINFO.Y_LO = Int(yBot + 0.5)
            XA = Int(xBot + 0.5)
        End If
        
    End If
    
    If dx = 0 Then
'        area = LMax(overT, yTop) 'er .. why did I have this
        If xBot < -0.5 Then
            m_RegionReset overT, yTop
        ElseIf xBot >= sWmp5 Then
            m_RegionReset overT, yTop, True
        End If
    Else
        dx = xTop - xBot
    End If
    
    dy = yTop - yBot
    
    YA = LMax(Int(yBot + 0.5), 0)
    XA = LMax(Int(xBot + 0.5), 0)
    Y_ED = LMin(Int(yTop + 0.5), SH_Top)
    X_ED = LMin(Int(xTop + 0.5), SW_Right)
    X_ST = LMin(XA, SW_Right)
    build_length = YA - 0.5 - yBot
    xi_lo = xBot + dx * build_length / dy

End Sub
Private Function m_RegionReset(sngLo!, sngHi!, Optional ByVal ScreenRight_ As Boolean = False)
Dim YA_ As Long

    If m_IsRight Then
        If ScreenRight_ Then
            For YA_ = LMax(sngLo, 0) To LMin(sngHi, SH_Top)
                TriBits(YA_).R1 = SW
                TriBits(YA_).R2 = SW_Right
            Next
        Else
            For YA_ = LMax(sngLo, 0) To LMin(sngHi, SH_Top)
                TriBits(YA_).R1 = 0
                TriBits(YA_).R2 = -1
            Next
        End If
    Else
        If ScreenRight_ Then
            For YA_ = LMax(sngLo, 0) To LMin(sngHi, SH_Top)
                TriBits(YA_).L1 = SW
                TriBits(YA_).L2 = SW_Right
            Next
        Else
            For YA_ = LMax(sngLo, 0) To LMin(sngHi, SH_Top)
                TriBits(YA_).L1 = 0
                TriBits(YA_).L2 = -1
            Next
        End If
    End If
    
End Function

Private Function LMax(ByVal sVar1 As Single, ByVal sVar2 As Single)
    If sVar1 < sVar2 Then
        LMax = Int(sVar2 + 0.5)
    Else
        LMax = Int(sVar1 + 0.5)
    End If
End Function
Private Function LMin(ByVal sVar1 As Single, ByVal sVar2 As Single)
    If sVar1 < sVar2 Then
        LMin = Int(sVar1 + 0.5)
    Else
        LMin = Int(sVar2 + 0.5)
    End If
End Function

Public Sub RasterizeMatrix(ColorMatrix As ColorMatrix)

End Sub
Private Sub m_Edge(P1 As FP_NODE, P2 As FP_NODE, ColorL As Long, ColorR As Long)

End Sub
Private Sub m_ClipperSteepNegL()

End Sub
Private Sub m_ClipperShallowPosL()

End Sub
Private Sub m_ClipperSteepPosL()

End Sub

Private Sub RenderExample(sx1 As Single, sy1 As Single, sx2 As Single, sy2 As Single, sx3 As Single, sy3 As Single)
Dim PicY       As Long
Dim LngY       As Long
Dim LngX       As Long
Dim X1         As Long
Dim X2         As Long
Dim R1         As Long
Dim R2         As Long
Dim Alpha1     As Long
Dim GrayScale1 As Long
Dim sAlpha2    As Single

    'First, some custom variables just for this example sub ..
    
    GrayScale1 = 1 + 256 + 65536
    Alpha1 = 255
    
    'Now, two things must happen:
    
    '1. (try Form_Resize) Call TriSurfaceInitialize(SurfaceWidth, SurfaceHeight)
    
    '2.  when you are ready to render ..
    CalculateRegions sx1, sy1, sx2, sy2, sx3, sy3
    
    'And this is how it works .. by the way, copy this sub to your form
    'or wherever ..
    
    'required data:
    
              'TriSurfaceINFO .. a type
    For LngY = TriSurfaceINFO.Y_LO To TriSurfaceINFO.Y_HI
    
            'TriBits(    ) .. type array
        X1 = TriBits(LngY).L1
        X2 = TriBits(LngY).L2
        R1 = TriBits(LngY).R1
        R2 = TriBits(LngY).R2
        PicY = TriSurfaceINFO.HighM - LngY
        
        For LngX = X1 To X2
                     'triAlpha(          ) .. float array
            sAlpha2 = triAlpha(LngX, LngY)
'            PSet (LngX, PicY), Int(sAlpha2 * Alpha1 + 0.5) * GrayScale1
        Next
        
'        Line (X2 + 1, PicY)-(R1, PicY), Alpha1 * GrayScale1
        
        For LngX = R1 To R2
            sAlpha2 = triAlpha(LngX, LngY)
'            PSet (LngX, PicY), Int(sAlpha2 * Alpha1 + 0.5) * GrayScale1
        Next
     
    Next

End Sub

Private Sub RenderExampleHARDCORE(LngAry1() As Long, sx1 As Single, sy1 As Single, sx2 As Single, sy2 As Single, sx3 As Single, sy3 As Single)
Dim LngY       As Long
Dim LngX       As Long
Dim L1D        As Long
Dim X1         As Long
Dim X2         As Long
Dim R1         As Long
Dim R2         As Long
Dim Alpha1     As Long
Dim GrayScale1 As Long
Dim sAlpha2    As Single

    'First, some custom variables just for this example sub ..
    
    GrayScale1 = 1 + 256 + 65536
    Alpha1 = 255
    
    'Now, two things must happen:
    
    '1. (try Form_Resize) Call TriSurfaceInitialize(SurfaceWidth, SurfaceHeight)
    
    '2.  when you are ready to render ..
    CalculateRegions sx1, sy1, sx2, sy2, sx3, sy3
    
    'Finally, there are three important names to recognize:
    'TriSurfaceINFO - a type
    'TriBits()      - a type array
    'triAlpha()     - a float array
    
    'And this is how it works .. by the way, copy this sub to your form
    'or wherever ..
    
    For LngY = TriSurfaceINFO.Y_LO To TriSurfaceINFO.Y_HI
    
        X1 = TriBits(LngY).L1
        X2 = TriBits(LngY).L2
        R1 = TriBits(LngY).R1
        R2 = TriBits(LngY).R2
        
        L1D = LngY * TriSurfaceINFO.Wide
        
        For LngX = X1 To X2
            sAlpha2 = triAlpha(LngX, LngY)
            LngAry1(LngX + L1D) = Int(sAlpha2 * Alpha1 + 0.5) * GrayScale1
        Next
        
        For LngX = X2 + 1 To R1 - 1
            LngAry1(LngX + L1D) = Alpha1 * GrayScale1
        Next
        
        For LngX = R1 To R2
            sAlpha2 = triAlpha(LngX, LngY)
            LngAry1(LngX + L1D) = Int(sAlpha2 * Alpha1 + 0.5) * GrayScale1
        Next
     
    Next

End Sub


