VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'smooth triangle mesh preview by dafhi

'a highly memory-inefficient method for
'rendering smooth triangles as a 'mesh'.

'This system uses roughly 7x32 bits per pixel.

'There are 7 arrays dimensioned to match the viewport.

'1. alpha surface for one triangle. 32 bit float
'2. Four 32 bit float arrays
' A. sRed() for red accumulation of multiple triangles in a pixel.
' B. sGrn()
' C. sBlu()
' D. sAlp()
'3. A stack (Longs) that holds pixel location for 'quick erase'
' of only data at triangle edge
'4. The regular draw surface

'Non P-Code Project1.exe sometimes gets overflow error.
'I use On Error for a relatively safe (and mostly unexpected)
'process termination.

Dim FormDib As SurfaceWrapper

Private Type Point3D
 sx  As Single
 sy  As Single
 sZ  As Single
End Type

Dim I As Long
Dim J As Long

Dim PixelSize As Long
Dim SizeM As Long

Dim SCW As Long
Dim SCH As Long

Dim Tick&
Dim TickPrev&
Dim Frame&
Dim FrameReset&

Private Type Build_Primitive
  Points()   As Point3D
  PointRef() As Long
  Colors()   As Long
  PointCount As Long
  Build_Type As Long
  sSize      As Single
  sScale     As Single
  scaleRot   As Single
  scaleRoti  As Single
End Type

Dim PrimA As Build_Primitive

Dim rotation_speed!

Private Const BUILD_RADIAL As Long = 1
Private Const BUILD_RIBBON As Long = 2

Private Sub Form_Load()

    ScaleMode = vbPixels
    
    PrimA.Build_Type = BUILD_RADIAL 'test points only set up for this mode
    
    PixelSize = 2
    
    FrameReset = 100

    BackColor = vbBlue
    
    Caption = "smooth triangles (test render)"
    
    Move 0, 0, 350 * Screen.TwipsPerPixelX, 267 * Screen.TwipsPerPixelY
    Show
    
    CreateRadialFan PrimA, 16
    
    TickPrev = timeGetTime
    
    PrimA.sScale = -Cos(PrimA.scaleRot)
    PrimA.scaleRot = PrimA.scaleRot + PrimA.scaleRoti
    PrimA.scaleRot = PrimA.scaleRot - TwoPi * Int(PrimA.scaleRot / TwoPi)
    
    On Local Error GoTo OHNO
    
    Do While DoEvents
        
        RenderExampleHARDCORE FormDib.Dib32, PrimA, BackColor, SCW / 2, SCH / 2
        
        StretchDIBits hDC, 0, 0, SCW * PixelSize, SCH * PixelSize, 0, 0, SCW, SCH, FormDib.Dib32(0), FormDib.BMI.bmiHeader, 0, vbSrcCopy

        Rot8 PrimA, rotation_speed
        
        Frame = Frame + 1
        If Frame = FrameReset Then
            Tick = timeGetTime
            Caption = Round(1000 * FrameReset / LMax((Tick - TickPrev), 1), 1) & " fps - try Up/Down arrows or (Shift) L/R"
            TickPrev = Tick
            Frame = 0
            Sleep 1
        End If
        
    Loop
    
OHNO:
    
    ClearSurface FormDib

End Sub
Private Sub SetPixelSize(Optional ByVal Size_ As Long)
    
    If Size_ > 0 Then
        PixelSize = Size_
    End If
    
    If PixelSize < 1 Then PixelSize = 1
    If PixelSize > 50 Then PixelSize = 50
    
    SizeM = PixelSize - 1
    
    ResizeCalcSurfaces

End Sub

Private Sub RenderExampleHARDCORE(LngAry1() As Long, Prim3D As Build_Primitive, Optional ByVal BackColor_& = vbWhite, Optional ByVal offX_ As Single = 50, Optional ByVal offY_ As Single = 50)
Dim LngY       As Long
Dim LngX       As Long
Dim L1D        As Long
Dim X1         As Long
Dim X2         As Long
Dim R1         As Long
Dim R2         As Long
Dim Alpha1     As Long
Dim sAlpha2    As Single
Dim sAlpha1    As Single
Dim LBlu       As Long
Dim LGrn       As Long
Dim LRed       As Long
Dim TriLoop    As Long
Dim L1DY       As Long
Dim sRd        As Single
Dim sGr        As Single
Dim sBl        As Single
Dim RedD       As Long
Dim GrnD       As Long
Dim BluD       As Long
Dim ForeColor_ As Long
Dim PT1_       As Long
Dim PT2_       As Long
Dim PT3_       As Long

    LLPtr = -1
    
    Alpha1 = 255
    
    PT1_ = 1
    PT2_ = 2
    PT3_ = 3
    
    For TriLoop = 3 To Prim3D.PointCount
        If TriLoop = 16 Then
            X1 = X1
        End If
    
        CalculateRegions Prim3D.Points(PT1_).sx * Prim3D.sScale + offX_, Prim3D.Points(PT1_).sy * Prim3D.sScale + offY_, _
         Prim3D.Points(PT2_).sx * Prim3D.sScale + offX_, Prim3D.Points(PT2_).sy * Prim3D.sScale + offY_, _
         Prim3D.Points(PT3_).sx * Prim3D.sScale + offX_, Prim3D.Points(PT3_).sy * Prim3D.sScale + offY_
         
        If Prim3D.Build_Type <> BUILD_RADIAL Then
            PT1_ = PT2_
        End If
        
        ForeColor_ = Prim3D.Colors(PT3_)
        
        LRed = ((ForeColor_ And &HFF0000) \ 256&) \ 256&
        LGrn = (ForeColor_ And &HFF00&) \ 256&
        LBlu = ForeColor_ And &HFF&
        
        sRd = LRed / 255&
        sGr = LGrn / 255&
        sBl = LBlu / 255&
        
        PT2_ = PT3_
        PT3_ = PT3_ + 1
        
        LBlu = ForeColor_ And &HFF&
        ForeColor_ = (ForeColor_ And &HFF00&) + 256& * (LBlu * 256&) + (ForeColor_ \ 256&) \ 256&
        
        For LngY = TriSurfaceINFO.Y_LO To TriSurfaceINFO.Y_HI
        
            X1 = TriBits(LngY).L1
            X2 = TriBits(LngY).L2
            R1 = TriBits(LngY).R1
            R2 = TriBits(LngY).R2
            
            L1DY = LngY * TriSurfaceINFO.Wide
            
            For LngX = X1 To X2
            
                sAlpha1 = triAlpha(LngX, LngY)
                
                L1D = L1DY + LngX
                If sAlpha1 > 0 Then
                
                    If sAlp(L1D) = 0 Then
                        LLPtr = LLPtr + 1  'a stack system records only pixel
                        LPtrA(LLPtr) = L1D 'locations at triangle edge, and
                                           'this information is used for erase
                                           'at the last loop of this sub
                    End If
                    
                    sAlpha2 = sAlpha1 * sRd + sRed(L1D)
                    sRed(L1D) = sAlpha2
                    
                    sAlpha2 = sAlpha1 * sGr + sGrn(L1D)
                    sGrn(L1D) = sAlpha2
                    
                    sAlpha2 = sAlpha1 * sBl + sBlu(L1D)
                    sBlu(L1D) = sAlpha2
                    
                    sAlp(L1D) = sAlpha1 + sAlp(L1D)
                    
                End If
                
            Next
            
            'If you want to fill ..
            
'            For LngX = X2 + 1 To R1 - 1
'                LngAry1(LngX + L1DY) = ForeColor_
'            Next

            For LngX = R1 To R2
                
                sAlpha1 = triAlpha(LngX, LngY)
                
                L1D = L1DY + LngX
                If sAlpha1 > 0 Then
                
                    If sAlp(L1D) = 0 Then
                        LLPtr = LLPtr + 1
                        LPtrA(LLPtr) = L1D
                    End If
                    
                    sAlpha2 = sAlpha1 * sRd + sRed(L1D)
                    sRed(L1D) = sAlpha2
                    
                    sAlpha2 = sAlpha1 * sGr + sGrn(L1D)
                    sGrn(L1D) = sAlpha2
                    
                    sAlpha2 = sAlpha1 * sBl + sBlu(L1D)
                    sBlu(L1D) = sAlpha2
                    
                    sAlp(L1D) = sAlpha1 + sAlp(L1D)
                    
                End If

            Next
         
        Next 'LngY
        
    Next 'TriLoop
    
    LBlu = BackColor_ And &HFF&
    BackColor_ = (BackColor_ And &HFF00&) + 256& * (LBlu * 256&) + (BackColor_ \ 256&) \ 256&
    
    LBlu = ((BackColor_ And &HFF0000) / 256&) / 256&
    LGrn = (BackColor_ And &HFF00&) / 256&
    LRed = (BackColor_ And &HFF&)
    
    For LngX = 0 To LLPtr
        
        L1D = LPtrA(LngX)
        
        sAlpha1 = sAlp(L1D)
        sRd = sRed(L1D)
        sGr = sGrn(L1D)
        sBl = sBlu(L1D)
        
        If sRd > 1 Then sRd = 1
        If sGr > 1 Then sGr = 1
        If sBl > 1 Then sBl = 1
        If sAlpha1 > 1 Then sAlpha1 = 1
        sAlpha1 = 1 - sAlpha1
        
        LngAry1(L1D) = LRed * sAlpha1 + sRd * 255& + 256& * CLng(LGrn * sAlpha1 + sGr * 255& + 256& * CLng(LBlu * sAlpha1 + sBl * 255&))
    
        sRed(L1D) = 0 'erase alpha information
        sGrn(L1D) = 0
        sBlu(L1D) = 0
        sAlp(L1D) = 0
    
    Next
    
End Sub

Private Sub SetPoint(Prim3D As Build_Primitive, sx!, sy!, Optional Color_ As Long)
    Prim3D.PointCount = Prim3D.PointCount + 1
    ReDim Preserve Prim3D.Points(1 To Prim3D.PointCount)
    Prim3D.Points(Prim3D.PointCount).sx = sx * PrimA.sSize
    Prim3D.Points(Prim3D.PointCount).sy = sy * PrimA.sSize
    If Prim3D.PointCount > 2 Then
        ReDim Preserve Prim3D.Colors(3 To Prim3D.PointCount)
        Prim3D.Colors(Prim3D.PointCount) = Color_
    End If
End Sub
Private Sub Rot8(Prim3D As Build_Primitive, angle_ As Single)
Dim J1   As Long
Dim tmpX As Single
Dim tmpY As Single
    
    For J1 = 1 To Prim3D.PointCount
        With Prim3D.Points(J1)
            tmpX = .sx
            .sx = .sx * Cos(angle_) - Sin(angle_) * .sy
            .sy = .sy * Cos(angle_) + Sin(angle_) * tmpX
        End With
    Next

End Sub






Private Sub Form_Resize()
    
    SetPixelSize
    
End Sub
Private Sub ResizeCalcSurfaces()
    
    SCW = ScaleWidth / LMax(PixelSize, 1)
    SCH = ScaleHeight / LMax(PixelSize, 1)
    
    'FIRST REQUIREMENT for basic triangle element
    TriSurfaceInitialize SCW, SCH
    
    CreateSurface FormDib, SCW, SCH
    
    Flood FormDib, BackColor, , True
    
    'SECOND REQUIREMENT for 'mesh'
    Resize_sPlane sRed, FormDib.UB32
    Resize_sPlane sGrn, FormDib.UB32
    Resize_sPlane sBlu, FormDib.UB32
    Resize_sPlane sAlp, FormDib.UB32
    Resize_LPlane LPtrA, FormDib.UB32, LLPtr
    
    'set up a test mesh
    PrimA.sSize = Sqr(SCW * SCH) / 2.458
    CreateRadialFan PrimA
    
    rotation_speed = 0.001 * PixelSize
    
End Sub
Private Sub CreateRadialFan(Prim3D As Build_Primitive, Optional AddPointCount As Long = 0)
Dim LCountTMP As Long
    
    LCountTMP = (Prim3D.PointCount - 1) + AddPointCount
    
    Prim3D.PointCount = 0
    SetPoint PrimA, 0, 0
    
    If LCountTMP < 3 Then LCountTMP = 3
    For I = 1 To LCountTMP
        SetPoint PrimA, Cos(TwoPi * I / LCountTMP), Sin(TwoPi * I / LCountTMP), ARGBHSV(1530 * I / LCountTMP, 0.5 * (1 + Rnd), 128 + Rnd * 127)
    Next

End Sub
Private Sub Resize_sPlane(sAry_() As Single, L1D&)
    If L1D > -1 Then
        Erase sAry_
        ReDim sAry_(L1D)
    End If
End Sub
Private Sub Resize_LPlane(LAry_() As Long, ByVal L1D&, LLPtr_&)
    If L1D > -1 Then
        Erase LAry_
        ReDim LAry_(L1D)
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    Case vbKeyDown
        SetPixelSize PixelSize - 1
    Case vbKeyUp
        SetPixelSize PixelSize + 1
    Case vbKeyLeft, vbKeyL
        CreateRadialFan PrimA, -5 - 4 * (Shift <> 0)
    Case vbKeyRight, vbKeyR
        CreateRadialFan PrimA, 5 + 4 * (Shift <> 0)
    End Select
End Sub

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

