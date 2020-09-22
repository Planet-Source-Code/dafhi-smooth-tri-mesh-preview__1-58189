Attribute VB_Name = "mSurfaceWrapper"
Option Explicit

' mSurfaceWrapper 1.01a by dafhi

'If you see "Type not defined" error, add OLE Automation
'from the Visual Basic menu:  Project, References, OLE Automation (add to list)

'+-------------+--------------------------------------+
'| Description | Development of DMA graphics routines |
'+-------------+ for eventual release of a super      |
'| speedy sprite engine                               |
'+----------------------------------------------------+

'Advantage to using SurfaceWrapper:
'1.  Direct Memory Access (DMA)
'2.  Compatibility

'Disadvantages:
'1.  1d array makes it difficult, but I've worked out some
'details.  Have a peek inside Blit()
'2.  No Device Context association (this could be a good thing)
'3.  Gotta clear all Surfaces with ClearSurface before program exit.
'(memory leak if you don't)

'Blit() should keep you busy.


' === CopyMemoryMMX class ===

'Also worthy of recognition is the possibility of using the
'cMemory class which uses MMX or SSE if present.

'If you have the class you can uncomment the following line,
'Public cMemA As cMemory 'CopyMem / MMX / SSE

' and look for LINE PAIRS as shown here, applying / unapplying comment mark:
'        cMemA.CopyMemory ByVal VarPtr(Surf.Dib32(DY)), ByVal VarPtr(Surf.Dib32(0)), Surf.BM.bmWidthBytes
'        CopyMemory Surf.Dib32(DY), Surf.Dib32(0), Surf.BM.bmWidthBytes


' ====== OTHER STUFF ======


'There are 'undocumented' subs like BlitSolid(),
'or undocumented and obscure BlitReverse() ..


'-------Notes about my 24 bit surface technique------

'24-bit bmps have what are called pad bytes if
'3 bytes per pixel * N pixels comes out to a non-
'multiple of four.

'Like if you had a 3x3 bitmap, Win32 will arrange
'the byte array like this

'+---+---+---+
'|BGR|BGR|BGR|xxx '12 bytes per scanline
'+---+---+---+
'|BGR|BGR|BGR|xxx
'+---+---+---+
'|BGR|BGR|BGR|xxx
'+---+---+---+

'What I do is create a 'Scanline24' type consisting of
'a SafeArray struct and an RGBTriple array which handles
'every normal pixel for each scanline.

'That's right.  A pointer is created for each scanline.
'The SafeArray's LBound is set like this

'+---+---+---+
'|6  |   |   |
'+---+---+---+
'|3  |   |   |
'+---+---+---+
'|0  |   |   |
'+---+---+---+

'So in this way, you will see me using this code
'to access a pixel,

'Surf.ScanLine24(Y).Row(Y*Surf.Bm.BmWidth + X).Red = Whatever

'Of course I could Send the ScanLine24 to a wrapper sub
'where this could happen:

'For X = SL24.SA.LBound To SL24.SA.LBound + 9 'Write 10 pixels
' SL24.Row(X).Red = Whatever
'Next

'ScanLine24s are not created if the surface is 32bpp (default).
'There is a speed advantage to Form1.Picture = HookPicture24()
'if you don't plan on having a lot of sprite area or alpha-
'blending.

'Overall, I reccommend using 32 bit surfaces.

'How to use:

''(Paste these lines in Form1)

'Dim DibMe As SurfaceWrapper 'access to Form.Picture
'Dim DibPic as SurfaceWrapper

'Private Sub Form_Resize()
' Me.Picture = HookPicture32(DibMe, ScaleWidth, ScaleHeight)
'End Sub

'Private Sub Form_Load()
' ScaleMode = vbPixels
' CreateFromFile DibPic, "object.bmp", , vbBlack 'make a bmp with black and non-black pixels
'End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button Then
'        Blit DibMe, X, Y, DibPic
'        Refresh
'    End If
'End Sub

'Private Sub Form_Unload()
'    ClearSurface DibMe
'End Sub

Type RGBTriple
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Type RGBQUAD
 Blue  As Byte
 Green As Byte
 Red   As Byte
 Alpha As Byte
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

'Private Type SAFEARRAYBOUND
'    cElements As Long
'    lLbound As Long
'End Type

'Private Type SAFEARRAY2D
'    cDims As Integer
'    fFeatures As Integer
'    cbElements As Long
'    cLocks As Long
'    pvData As Long
'    Bounds(1) As SAFEARRAYBOUND
'End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    reserved As Long
End Type

Private Type ScanSeg
    X1 As Long
    X2 As Long
End Type

Public Type MaskInfo
    PartsM As Long
    Red As Byte
    Green As Byte
    Blue As Byte
    Rsv As Byte
    DrawPart() As ScanSeg
End Type

Private Type ScanLine
    Row() As RGBTriple
    SA As SAFEARRAY1D
End Type

Private Type BppInfo
    PixelTopLeft As Long
    BlueTopLeft As Long
    BlueTopRight As Long
    BlueRight As Long
    BytesPixel As Long
End Type

Public Type SurfaceWrapper
 MaskInfo() As MaskInfo
 BMI As BITMAPINFO
 BM As BITMAP
 tSA As SAFEARRAY1D
 hDib As Long
 TotalPixels As Long
 WideM As Long
 HighM As Long
 UsingMask As Boolean
 UB32 As Long
 Proc As BppInfo
 Dib32() As Long
 Dib24() As Byte
 ScanLine24() As ScanLine
End Type

Public Const DIB_RGB_COLORS As Long = 0
Public Const GrayScaleRGB As Long = 1 + 256& + 65536

'store values for use by nested subs
Public DrawWidthM As Long
Private SrcLeft As Long
Public SrcTop As Long
Private SrcBot As Long
Public SrcTopLeft As Long
Public DestBot As Long
Public DestTopLeft As Long

'General Purpose
Private IA As Long
Private DWBytes As Long
Private DrawX As Long
Private DrawY As Long
Private PicBmA As PicBmp
Private strFileFolder As String

Declare Function StretchDIBits Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal dx As Long, _
         ByVal dy As Long, _
         ByVal SrcX As Long, _
         ByVal SrcY As Long, _
         ByVal wSrcWidth As Long, _
         ByVal wSrcHeight As Long, _
         lpBits As Any, _
         lpBitsInfo As BITMAPINFOHEADER, _
         ByVal wUsage As Long, _
         ByVal dwRop As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC&, pBitmapInfo As BITMAPINFO, ByVal un&, lplpVoid&, ByVal handle&, ByVal dw&) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, IPic As IPicture) As Long
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Const BI_RGB = 0&

Dim RGBTriA  As RGBTriple
Dim RGBTriB  As RGBTriple
Dim RGBTriC  As RGBTriple
Dim RGBTriD As RGBTriple

Dim FGColor&
Dim BGColor&

Dim FGGrn&
Dim FGBlu&
Dim FGRed&

'ARGBHSV() Function
Private BGBlu&
Private BGGrn&
Private BGRed&
Dim subt!

Dim m_alpha!

Public Function HookPicture32(Surf As SurfaceWrapper, ByVal nWidth&, ByVal nHeight&) As Picture
    m_PicCommon Surf, nWidth, nHeight, 32, HookPicture32
End Function
Public Function HookPicture24(Surf As SurfaceWrapper, ByVal nWidth&, ByVal nHeight&) As Picture
    m_PicCommon Surf, nWidth, nHeight, 24, HookPicture24
End Function

Public Sub CreateFromFile(Surf As SurfaceWrapper, strFileName$, Optional ByVal MaskColor = -1, Optional ByVal strFolder$ = "", Optional Force32 As Boolean = True)
Dim tBM As BITMAP, sPic As StdPicture
Dim CDC&, I1&, RootDirFound As Boolean
Dim bBits() As Byte

    m_Adjust_strFileFolder strFolder
    
    For I1 = 1 To Len(strFileName)
        If Mid$(strFileName, I1, 1) = ":" Then
            RootDirFound = True
            Exit For
        End If
    Next
    
    CreateSurface Surf, 10, 10
    Flood Surf, vbWhite, 1
    
On Local Error GoTo OHNO
    
    If RootDirFound Then
        Set sPic = LoadPicture(strFileName)
    Else
        Set sPic = LoadPicture(strFileFolder & strFileName)
    End If
    
    CDC = CreateCompatibleDC(0)           ' Temporary device
    DeleteObject SelectObject(CDC, sPic)  ' Converted bitmap
    
    GetObjectAPI sPic, Len(Surf.BM), Surf.BM
    
    If m_MakeSurf(Surf, Surf.BM.bmWidth, Surf.BM.bmHeight) Then
    
      If Force32 Then
        ReDim bBits(Surf.BM.bmWidthBytes * Surf.BM.bmHeight - 1)
        CopyMemory bBits(0), ByVal Surf.BM.bmBits, Surf.BM.bmWidthBytes * Surf.BM.bmHeight
        ReDim Surf.Dib32(Surf.UB32)
        m_CopyTo32 Surf, bBits
        Surf.BM.bmBitsPixel = 32
      Else
        ReDim Surf.Dib24(Surf.BM.bmWidthBytes * Surf.BM.bmHeight - 1)
        CopyMemory Surf.Dib24(0), ByVal Surf.BM.bmBits, Surf.BM.bmWidthBytes * Surf.BM.bmHeight
        For IA = 0 To Surf.Proc.BlueTopRight Step 3
            If Surf.Dib24(IA) <> 0 Then
            IA = IA
            End If
        Next
        m_SetScans24 Surf
      End If
      
      If MaskColor > -1 Then
        DefineMask Surf, MaskColor
      End If
      
    End If 'MakeSurf
 
OHNO:
 
    DeleteDC CDC
 
End Sub

'These subs are intended to work together, but calling DefineMask is not necessary.
Public Sub CreateSurface(Surf As SurfaceWrapper, ByVal Width%, ByVal Height%, Optional ByVal FillColor_&) ', Optional ByVal MaskColor_& = -1)
    Surf.BM.bmBitsPixel = 32
    m_MakeSurf Surf, Width, Height
    Surf.BM.bmWidthBytes = Surf.Proc.BytesPixel * Width
    Flood Surf, FillColor_
End Sub
Public Sub DefineMask(Surf As SurfaceWrapper, Optional ByVal MaskColor&)
Dim TL_&

    'This sub is called by CreateFromFile() if maskcolor parameter there > -1.
    
    'Another way you could use this sub:
    'Private Sub SomeSub()
    'Dim SurfA as SurfaceWrapper
    ' CreateSurface SurfA, 40, 40
    ' Flood SurfA, vbWhite,100 '100 random white pixels
    ' DefineMask SurfA, 0, 0, 0 'RGB(0,0,0) will define a mask
    'structure where only non-black pixels draw to destination.

    If Surf.TotalPixels < 1 Then Exit Sub
    
    Erase Surf.MaskInfo
    ReDim Surf.MaskInfo(Surf.HighM)
    
    For IA = 0 To Surf.HighM
        ReDim Surf.MaskInfo(IA).DrawPart(Int(Surf.BM.bmWidth / 2) + 1)
        m_SetMask Surf, Surf.MaskInfo(IA), -1, 0, MaskColor, TL_, IA
    Next
    
    Surf.UsingMask = True
    
End Sub

' = Pixel manipulation =
Sub PutPixel(Surf As SurfaceWrapper, ByVal X&, ByVal Y&, Color_&)
    If X > -1 And X < Surf.BM.bmWidth Then
        If Y > -1 And Y < Surf.BM.bmHeight Then
            Surf.Dib32(Y * Surf.BM.bmWidth + X) = Color_
        End If
    End If
End Sub
Sub Blit(ByVal X&, ByVal Y&, Dest1 As SurfaceWrapper, Src1 As SurfaceWrapper, Optional ByVal AlphaVal As Byte = 255)
Dim SrcX&, Y1&

    If RectsIntersect(Dest1, Src1, X, Y) Then
    'This is a wrapper that sets up vars for 1d array entry point,
    'taking rect clipping into consideration
    
        m_alpha = AlphaVal / 255
    
        If Src1.UsingMask Then
            If Dest1.BM.bmBitsPixel = 32 Then
                If Src1.BM.bmBitsPixel = 32 Then
                    If AlphaVal = 255 Then
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_MaskBlit Y, SrcTopLeft, SrcTopLeft + DrawWidthM, Src1.MaskInfo(SrcTop), Dest1.Dib32, Src1.Dib32
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                        Next
                    Else
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_MaskBlit_D32_S32_Alpha Y, m_alpha, SrcTopLeft, SrcTopLeft + DrawWidthM, 0, 0, Src1.MaskInfo(SrcTop), Dest1.Dib32, Src1.Dib32
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                        Next
                    End If
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                    If AlphaVal = 255 Then
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_MaskBlit_D32_S24 Y, m_alpha, SrcTopLeft, SrcTopLeft + DrawWidthM, 0, 0, Src1.MaskInfo(SrcTop), Dest1.Dib32, Src1.ScanLine24(SrcTop).Row
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                        Next
                    Else 'AlphaVal <> 255
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_MaskBlit_D32_S24_Alpha Y, m_alpha, SrcTopLeft, SrcTopLeft + DrawWidthM, 0, 0, Src1.MaskInfo(SrcTop), Dest1.Dib32, Src1.ScanLine24(SrcTop).Row
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                        Next
                    End If 'AlphaVal = 255
                End If
            ElseIf Dest1.BM.bmBitsPixel = 24 Then
                Y1 = Dest1.HighM - Y
                If Src1.BM.bmBitsPixel = 32 Then
                    If AlphaVal = 255 Then
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_MaskBlit_D24_S32 Y, SrcTopLeft, SrcTopLeft + DrawWidthM, 0, 0, 0, Src1.MaskInfo(SrcTop), Dest1.ScanLine24(Y1).Row, Src1.Dib32
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                            Y1 = Y1 - 1
                        Next
                    Else 'AlphaVal <> 255
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_MaskBlit_D24_S32_Alpha Y, SrcTopLeft, SrcTopLeft + DrawWidthM, 0, 0, 0, Src1.MaskInfo(SrcTop), Dest1.ScanLine24(Y1).Row, Src1.Dib32
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                            Y1 = Y1 - 1
                        Next
                    End If 'AlphaVal = 255
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                    If AlphaVal = 255 Then
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_MaskBlit_D24_S24 Y, SrcTopLeft, SrcTopLeft + DrawWidthM, 0, 0, 0, Src1.MaskInfo(SrcTop), Dest1.ScanLine24(Y1).Row, Src1.ScanLine24(SrcTop).Row
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                            Y1 = Y1 - 1
                        Next
                    Else 'AlphaVal <> 255
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_MaskBlit_D24_S24_Alpha Y, SrcTopLeft, SrcTopLeft + DrawWidthM, 0, 0, 0, Src1.MaskInfo(SrcTop), Dest1.ScanLine24(Y1).Row, Src1.ScanLine24(SrcTop).Row
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                            Y1 = Y1 - 1
                        Next
                    End If 'AlphaVal = 255
                End If
            End If
        Else 'not using mask
            If Dest1.BM.bmBitsPixel = 32 Then
                If Src1.BM.bmBitsPixel = 32 Then
                    If AlphaVal = 255 Then
                        DWBytes = DrawWidthM * 4 + 4
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            'if you don't have the MMX copymemory class, you can comment this line out and use the line below
'                            cMemA.CopyMemory ByVal VarPtr(Dest1.Dib32(Y)), ByVal VarPtr(Src1.Dib32(SrcTopLeft)), DWBytes

                            CopyMemory Dest1.Dib32(Y), Src1.Dib32(SrcTopLeft), DWBytes
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                        Next
                    Else 'AlphaVal <> 255
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            SrcX = SrcTopLeft
                            For X = Y To Y + DrawWidthM
                                BGColor = Dest1.Dib32(X) And &HFFFFFF
                                FGColor = Src1.Dib32(SrcX)
                                Dest1.Dib32(X) = BGColor + _
                                 (m_alpha * ((FGColor And &HFF0000) - (BGColor And &HFF0000)) And &HFF0000 Or _
                                 m_alpha * ((FGColor And &HFF00&) - (BGColor And &HFF00&)) And &HFF00& Or _
                                 m_alpha * ((FGColor And &HFF&) - (BGColor And &HFF&)) And &HFF&)
                                SrcX = SrcX + 1
                            Next
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                        Next
                    End If 'AlphaVal = 255
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                    If AlphaVal = 255 Then
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_ScanNormal24To32 Dest1.Dib32, Src1.ScanLine24(SrcTop).Row, Y, Y + DrawWidthM, SrcTopLeft
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                        Next
                    Else 'AlphaVal <> 255
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_ScanNormal24To32_Alpha Dest1.Dib32, Src1.ScanLine24(SrcTop).Row, Y, Y + DrawWidthM, SrcTopLeft
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                        Next
                    End If 'AlphaVal = 255
                End If
            ElseIf Dest1.BM.bmBitsPixel = 24 Then 'reminder: Dest bpp 32, src non-mask above section
                Y1 = Dest1.HighM - Y
                If Src1.BM.bmBitsPixel = 24 Then
                    If AlphaVal = 255 Then
                        DWBytes = DrawWidthM * 3 + 3
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            CopyMemory Dest1.ScanLine24(Y1).Row(Y).Blue, Src1.ScanLine24(SrcTop).Row(SrcTopLeft).Blue, DWBytes
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                            Y1 = Y1 - 1
                        Next
                    Else 'AlphaVal <> 255
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_ScanNormal24To24_Alpha Dest1.ScanLine24(Y1).Row, Src1.ScanLine24(SrcTop).Row, Y, Y + DrawWidthM, SrcTopLeft
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                            Y1 = Y1 - 1
                        Next
                    End If 'AlphaVal = 255
                ElseIf Src1.BM.bmBitsPixel = 32 Then
                    If AlphaVal = 255 Then
                        DWBytes = DrawWidthM * 3 + 3
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_ScanNormal32To24 Dest1.ScanLine24(Y1).Row, Src1.Dib32, Y, Y + DrawWidthM, SrcTopLeft
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                            Y1 = Y1 - 1
                        Next
                    Else 'AlphaVal <> 255
                        For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                            m_ScanNormal32To24_Alpha Dest1.ScanLine24(Y1).Row, Src1.Dib32, Y, Y + DrawWidthM, SrcTopLeft
                            SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                            SrcTop = SrcTop - 1
                            Y1 = Y1 - 1
                        Next
                    End If 'AlphaVal = 255
                End If 'Src1.BM.bmBitsPixel = 24
            End If 'Dest1.BM.bmBitsPixel = 24
        End If
    End If
End Sub
Private Sub m_ScanNormal24To24_Alpha(Dest24() As RGBTriple, Src24() As RGBTriple, ByVal DestX_Start&, DestX_End&, ByVal SrcTopLeft1&)
    For DestX_Start = DestX_Start To DestX_End
        RGBTriB = Dest24(DestX_Start)
        Dest24(DestX_Start).Blue = RGBTriB.Blue + m_alpha * (Src24(SrcTopLeft1).Blue - CLng(RGBTriB.Blue))
        Dest24(DestX_Start).Green = RGBTriB.Green + m_alpha * (Src24(SrcTopLeft1).Green - CLng(RGBTriB.Green))
        Dest24(DestX_Start).Red = RGBTriB.Red + m_alpha * (Src24(SrcTopLeft1).Red - CLng(RGBTriB.Red))
        SrcTopLeft1 = SrcTopLeft1 + 1
    Next
End Sub
Private Sub m_ScanNormal32To24(Dest24() As RGBTriple, Src32() As Long, ByVal DestX_Start&, DestX_End&, ByVal SrcTopLeft1&)
    For DestX_Start = DestX_Start To DestX_End
        FGColor = Src32(SrcTopLeft1)
        Dest24(DestX_Start).Blue = FGColor And &HFF&
        Dest24(DestX_Start).Green = (FGColor And &HFF00&) / 256&
        Dest24(DestX_Start).Red = (FGColor And &HFF0000) / 65536
        SrcTopLeft1 = SrcTopLeft1 + 1
    Next
End Sub
Private Sub m_ScanNormal32To24_Alpha(Dest24() As RGBTriple, Src32() As Long, ByVal DestX_Start&, DestX_End&, ByVal SrcTopLeft1&)
    For DestX_Start = DestX_Start To DestX_End
        FGColor = Src32(SrcTopLeft1)
        RGBTriB = Dest24(DestX_Start)
        Dest24(DestX_Start).Blue = RGBTriB.Blue + m_alpha * ((FGColor And &HFF&) - CLng(RGBTriB.Blue))
        Dest24(DestX_Start).Green = RGBTriB.Green + m_alpha * ((FGColor And &HFF00&) / 256& - CLng(RGBTriB.Green))
        Dest24(DestX_Start).Red = RGBTriB.Red + m_alpha * ((FGColor And &HFF0000) / 65536 - CLng(RGBTriB.Red))
        SrcTopLeft1 = SrcTopLeft1 + 1
    Next
End Sub
Private Sub m_ScanNormal24To32(Dest32() As Long, Src24() As RGBTriple, ByVal DestX_Start&, DestX_End&, ByVal SrcTopLeft1&)
    For DestX_Start = DestX_Start To DestX_End
        Dest32(DestX_Start) = Src24(SrcTopLeft1).Blue Or _
         Src24(SrcTopLeft1).Green * 256& Or _
         Src24(SrcTopLeft1).Red * 65536
        SrcTopLeft1 = SrcTopLeft1 + 1
    Next
End Sub
Private Sub m_ScanNormal24To32_Alpha(Dest32() As Long, Src24() As RGBTriple, ByVal DestX_Start&, DestX_End&, ByVal SrcTopLeft1&)
    For DestX_Start = DestX_Start To DestX_End
        BGColor = Dest32(DestX_Start) And &HFFFFFF
        Dest32(DestX_Start) = BGColor + _
         (m_alpha * (Src24(SrcTopLeft1).Red * 65536 - (BGColor And &HFF0000)) And &HFF0000 Or _
         m_alpha * (Src24(SrcTopLeft1).Green * 256& - (BGColor And &HFF00&)) And &HFF00& Or _
         m_alpha * (Src24(SrcTopLeft1).Blue - (BGColor And &HFF&)) And &HFF&)
        SrcTopLeft1 = SrcTopLeft1 + 1
    Next
End Sub
Sub BlitSolid(ByVal X&, ByVal Y&, Dest1 As SurfaceWrapper, Src1 As SurfaceWrapper, Color&)
Dim SrcX&
    If RectsIntersect(Dest1, Src1, X, Y) Then
        If Src1.UsingMask Then
            If Dest1.BM.bmBitsPixel = 32 Then
                If Src1.BM.bmBitsPixel = 32 Then
                    For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                        m_MaskBlitSolid Y, SrcTopLeft, SrcTopLeft + DrawWidthM, 0, 0, 0, Src1.MaskInfo(SrcTop), Dest1, Src1, Color
                        SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                        SrcTop = SrcTop - 1
                    Next
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                    'Possible future implementation
                End If
            ElseIf Dest1.BM.bmBitsPixel = 24 Then
                'Possible future implementation
                If Src1.BM.bmBitsPixel = 32 Then
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                End If
            End If
        Else 'not using mask
            If Dest1.BM.bmBitsPixel = 32 Then
                If Src1.BM.bmBitsPixel = 32 Then
                    For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                        SrcX = SrcTopLeft
                        For X = Y To Y + DrawWidthM
                            Dest1.Dib32(X) = Color
                            SrcX = SrcX + 1
                        Next
                        SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                    Next
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                    'Possible future implementation
                End If
            ElseIf Dest1.BM.bmBitsPixel = 24 Then
                'Possible future implementation
                If Src1.BM.bmBitsPixel = 32 Then
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                End If
            End If
        End If
    End If
End Sub

'Destination pixels are copied onto Source
Sub BlitReverse(Dest1 As SurfaceWrapper, ByVal X&, ByVal Y&, Src1 As SurfaceWrapper)
Dim SrcX&
    If RectsIntersect(Dest1, Src1, X, Y) Then
        If Src1.UsingMask Then
            If Dest1.BM.bmBitsPixel = 32 Then
                If Src1.BM.bmBitsPixel = 32 Then
                    For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                        m_MaskBlitReverse Y, SrcTopLeft, SrcTopLeft + DrawWidthM, 0, 0, 0, Src1.MaskInfo(SrcTop), Dest1, Src1
                        SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                        SrcTop = SrcTop - 1
                    Next
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                    'Possible future implementation
                End If
            ElseIf Dest1.BM.bmBitsPixel = 24 Then
                'Possible future implementation
                If Src1.BM.bmBitsPixel = 32 Then
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                End If
            End If
        Else 'not using mask
            If Dest1.BM.bmBitsPixel = 32 Then
                If Src1.BM.bmBitsPixel = 32 Then
                    For Y = DestTopLeft To DestBot Step -Dest1.BM.bmWidth
                        SrcX = SrcTopLeft
                        For X = Y To Y + DrawWidthM
                            Src1.Dib32(SrcX) = Dest1.Dib32(X)
                            SrcX = SrcX + 1
                        Next
                        SrcTopLeft = SrcTopLeft - Src1.BM.bmWidth
                    Next
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                    'Possible future implementation
                End If
            ElseIf Dest1.BM.bmBitsPixel = 24 Then
                'Possible future implementation
                If Src1.BM.bmBitsPixel = 32 Then
                ElseIf Src1.BM.bmBitsPixel = 24 Then
                End If
            End If
        End If
    End If
End Sub
Public Sub Flood(Surf As SurfaceWrapper, Optional ByVal Color&, Optional ByVal HowManyRandom_ As Long = 0, Optional FlipRedBlue As Boolean = False)
Dim dx&
Dim dy&
Dim RGBT As RGBTriple

    If FlipRedBlue Then Color = FlipRB(Color)

    If Surf.TotalPixels < 1 Then Exit Sub
    
    If Surf.BM.bmBitsPixel = 32 Then
        If HowManyRandom_ Then
            For dx = 1 To HowManyRandom_
                Surf.Dib32(Rnd * Surf.UB32) = Color
            Next
        Else
            For dx = 0 To Surf.WideM
                Surf.Dib32(dx) = Color
            Next
            For dy = Surf.BM.bmWidth To Surf.Proc.PixelTopLeft Step Surf.BM.bmWidth
                'if you have the MMX copymemory class, you can comment this line out and use the line below
                CopyMemory Surf.Dib32(dy), Surf.Dib32(0), Surf.BM.bmWidthBytes
'                cMemA.CopyMemory ByVal VarPtr(Surf.Dib32(DY)), ByVal VarPtr(Surf.Dib32(0)), Surf.BM.bmWidthBytes
            Next
        End If
    ElseIf Surf.BM.bmBitsPixel = 24 Then
        RGBT.Blue = Color And &HFF&
        RGBT.Green = (Color And &HFF00&) / 256&
        RGBT.Red = (Color And &HFF0000) / 65536
        If HowManyRandom_ Then
            For IA = 1 To HowManyRandom_
                dy = Int(Rnd * Surf.BM.bmHeight)
                Surf.ScanLine24(dy).Row(dy * Surf.BM.bmWidth + Int(Rnd * Surf.BM.bmWidth)) = RGBT
            Next
        Else
            m_FillScanLine Surf.ScanLine24(0).Row, RGBT, Surf
            For dy = 1 To Surf.HighM
                'if you have the MMX copymemory class, you can comment this line out and use the line below
                CopyMemory Surf.ScanLine24(dy).Row(dy * Surf.BM.bmWidth).Blue, Surf.ScanLine24(0).Row(0).Blue, Surf.BM.bmWidthBytes
'                cMemA.CopyMemory ByVal VarPtr(Surf.ScanLine24(DY).Row(DY * Surf.BM.bmWidth).Blue), ByVal VarPtr(Surf.ScanLine24(0).Row(0).Blue), Surf.BM.bmWidthBytes
            Next
        End If
    End If 'BitsPixel = 32
    
End Sub
Public Function FlipRB(Color_ As Long) As Long
Dim LBlu As Long
Dim LGrn As Long
Dim LRed As Long
    LBlu = Color_ And &HFF&
    FlipRB = (Color_ And &HFF00&) + 256& * (LBlu * 256&) + (Color_ \ 256&) \ 256&
End Function
Private Sub m_FillScanLine(RGB_() As RGBTriple, ColorTRI As RGBTriple, Surf As SurfaceWrapper)
Dim dx&
    For dx = 0 To Surf.WideM
        RGB_(dx) = ColorTRI
    Next
End Sub
Public Sub ClearSurface(Surf As SurfaceWrapper)
    m_ClearSafeArrays Surf
    Erase Surf.Dib32
    IA = DeleteObject(Surf.hDib)
    Surf.TotalPixels = 0
    Surf.UsingMask = False
End Sub
Private Sub m_ClearSafeArrays(Surf As SurfaceWrapper)
    If Surf.BM.bmBitsPixel = 24 Then
        CopyMemory ByVal VarPtrArray(Surf.Dib24), 0&, 4
        If Surf.TotalPixels > 0 And Surf.UsingMask Then
            For IA = 0 To Surf.HighM
                CopyMemory ByVal VarPtrArray(Surf.ScanLine24(IA).Row), 0&, 4
                Erase Surf.ScanLine24(IA).Row
            Next
            Erase Surf.ScanLine24
        End If
    ElseIf Surf.BM.bmBitsPixel = 32 Then
        CopyMemory ByVal VarPtrArray(Surf.Dib32), 0&, 4
    End If
End Sub

Private Sub m_PicCommon(Surf As SurfaceWrapper, nWidth&, nHeight&, BitDepth&, Pic_ As Picture)
Dim IID_IDispatch As GUID
Dim TmpTot&
nWidth = Abs(nWidth)
nHeight = Abs(nHeight)
Surf.TotalPixels = 0
Surf.BM.bmWidth = 0
Surf.BM.bmHeight = 0
TmpTot = nWidth * nHeight
If TmpTot > 0 Then
With Surf.BMI.bmiHeader
.biSize = Len(Surf.BMI.bmiHeader)
.biWidth = nWidth
.biHeight = nHeight
.biPlanes = 1
.biBitCount = BitDepth
End With
PicBmA.hBmp = CreateDIBSection(0, Surf.BMI, 0, 0, 0, 0)
IID_IDispatch.Data1 = &H20400: IID_IDispatch.Data4(0) = &HC0: IID_IDispatch.Data4(7) = &H46
PicBmA.Size = Len(PicBmA)
PicBmA.Type = vbPicTypeBitmap
OleCreatePictureIndirect PicBmA, IID_IDispatch, 1, Pic_
If Pic_ Then
 m_ClearSafeArrays Surf
 With Surf
 GetObjectAPI PicBmA.hBmp, Len(.BM), .BM
 .tSA.cDims = 1
 .tSA.cbElements = 1 - 3 * (.BM.bmBitsPixel = 32)
 .tSA.cElements = .BM.bmHeight * .BM.bmWidthBytes / .tSA.cbElements
 .tSA.pvData = .BM.bmBits
 .TotalPixels = TmpTot
 End With
 m_SetProcDims Surf
 If Surf.BM.bmBitsPixel = 32 Then
  CopyMemory ByVal VarPtrArray(Surf.Dib32), VarPtr(Surf.tSA), 4
 ElseIf Surf.BM.bmBitsPixel = 24 Then
  CopyMemory ByVal VarPtrArray(Surf.Dib24), VarPtr(Surf.tSA), 4
  m_SetScans24 Surf
 End If
 'Flood Surf, 0
Else
 Set Pic_ = Nothing
End If 'CreatePicture
End If 'TmpTot
End Sub
Private Sub m_SetScans24(Surf As SurfaceWrapper)
  ReDim Surf.ScanLine24(Surf.HighM)
  For IA = 0 To Surf.HighM
   m_SetScanLine24 Surf.ScanLine24(IA), VarPtr(Surf.Dib24(IA * Surf.BM.bmWidthBytes)), Surf.BM, IA * Surf.BM.bmWidth
  Next
End Sub
Public Function RectsIntersect(Dest1 As SurfaceWrapper, Src1 As SurfaceWrapper, X&, Y&) As Boolean
Dim DestRight_ As Long

    RectsIntersect = False
    
    If Src1.TotalPixels < 1 Or Dest1.TotalPixels < 1 Then Exit Function
    
    If X < Dest1.BM.bmWidth Then
        If Y < Dest1.BM.bmHeight Then
            DestRight_ = X + Src1.WideM
            If DestRight_ > -1 Then
                DestBot = Y + Src1.HighM
                If DestBot > -1 Then
                    RectsIntersect = True
                    If X < 0 Then
                        SrcLeft = -X
                        X = 0
                    Else
                        SrcLeft = 0
                    End If
                    If Y < 0 Then
                        SrcTop = -Y
                        Y = 0
                    Else
                        SrcTop = 0
                    End If
                    If DestRight_ > Dest1.WideM Then
                        DestRight_ = Dest1.WideM
                    End If
                    If DestBot > Dest1.HighM Then
                        DestBot = Dest1.HighM
                    End If
                    SrcTop = Src1.HighM - SrcTop
                    SrcTopLeft = SrcTop * Src1.BM.bmWidth + SrcLeft
                    SrcBot = SrcTopLeft - (DestBot - Y) * Src1.BM.bmWidth
                    DestTopLeft = (Dest1.HighM - Y) * Dest1.BM.bmWidth + X
                    DestBot = (Dest1.HighM - DestBot) * Dest1.BM.bmWidth
                    DrawWidthM = DestRight_ - X
                End If 'TopA > 0
            End If 'DestRight > 0
        End If 'Y <= cDest.Height
    End If 'X <= cDest.Width

End Function
Private Sub m_MaskBlit(ByVal Dest1D&, ByVal SrcPos_&, StopRight_&, ML_ As MaskInfo, Dest1() As Long, Src1() As Long)
Dim Seg_&
Dim X&, XS_&, XE_&
    For Seg_ = 0 To ML_.PartsM
        XS_ = ML_.DrawPart(Seg_).X1
        XE_ = ML_.DrawPart(Seg_).X2
        If XE_ >= SrcPos_ Then
            If XS_ < SrcPos_ Then
                XS_ = SrcPos_
            ElseIf XS_ > SrcPos_ Then
                Dest1D = Dest1D + XS_ - SrcPos_
                SrcPos_ = XS_
            End If
            If XE_ >= StopRight_ Then
                If XS_ <= StopRight_ Then
'                    cMemA.CopyMemory ByVal VarPtr(Dest1(Dest1D)), ByVal VarPtr(Src1(XS_)), (StopRight_ - XS_ + 1) * 4
                    ''Can use this if no cMemA class present
                    CopyMemory Dest1(Dest1D), Src1(XS_), (StopRight_ - XS_ + 1) * 4
                End If
                
                Exit Sub
            End If
            XE_ = XE_ + 1
            If XS_ <= XE_ Then
'                cMemA.CopyMemory ByVal VarPtr(Dest1(Dest1D)), ByVal VarPtr(Src1(XS_)), (XE_ - XS_) * 4
                CopyMemory Dest1(Dest1D), Src1(XS_), (XE_ - XS_) * 4
            End If
            Dest1D = Dest1D + XE_ - XS_
            SrcPos_ = XE_
        End If
    Next
End Sub
Private Sub m_MaskBlit_D24_S24(ByVal Dest1D&, ByVal SrcPos_&, StopRight_&, X&, XS_&, XE_&, ML_ As MaskInfo, Dest24() As RGBTriple, Src24() As RGBTriple)
Dim Seg_&
    For Seg_ = 0 To ML_.PartsM
        XS_ = ML_.DrawPart(Seg_).X1
        XE_ = ML_.DrawPart(Seg_).X2
        If XE_ >= SrcPos_ Then
            If XS_ < SrcPos_ Then
                XS_ = SrcPos_
            ElseIf XS_ > SrcPos_ Then
                Dest1D = Dest1D + XS_ - SrcPos_
                SrcPos_ = XS_
            End If
            If XE_ >= StopRight_ Then
                If XS_ <= StopRight_ Then
                    CopyMemory Dest24(Dest1D).Blue, Src24(XS_).Blue, (StopRight_ - XS_ + 1) * 3
                End If
                Exit Sub
            End If
            XE_ = XE_ + 1
            If XS_ <= XE_ Then
                CopyMemory Dest24(Dest1D).Blue, Src24(XS_).Blue, (XE_ - XS_) * 3
            End If
            Dest1D = Dest1D + XE_ - XS_
            SrcPos_ = XE_
'            For X = XS_ To XE_
'                Dest24(Dest1D) = Src24(X)
'                Dest1D = Dest1D + 1
'            Next
'            SrcPos_ = X
        End If
    Next
End Sub
Private Sub m_MaskBlit_D24_S24_Alpha(ByVal Dest1D&, ByVal SrcPos_&, StopRight_&, X&, XS_&, XE_&, ML_ As MaskInfo, Dest24() As RGBTriple, Src24() As RGBTriple)
Dim Seg_&
    For Seg_ = 0 To ML_.PartsM
        XS_ = ML_.DrawPart(Seg_).X1
        XE_ = ML_.DrawPart(Seg_).X2
        If XE_ >= SrcPos_ Then
            If XS_ < SrcPos_ Then
                XS_ = SrcPos_
            ElseIf XS_ > SrcPos_ Then
                Dest1D = Dest1D + XS_ - SrcPos_
                SrcPos_ = XS_
            End If
            If XE_ >= StopRight_ Then
                For X = XS_ To StopRight_
                    RGBTriB = Dest24(Dest1D)
                    Dest24(Dest1D).Blue = RGBTriB.Blue + m_alpha * (Src24(X).Blue - CLng(RGBTriB.Blue))
                    Dest24(Dest1D).Green = RGBTriB.Green + m_alpha * (Src24(X).Green - CLng(RGBTriB.Green))
                    Dest24(Dest1D).Red = RGBTriB.Red + m_alpha * (Src24(X).Red - CLng(RGBTriB.Red))
                    Dest1D = Dest1D + 1
                Next
                Exit Sub
            End If
            For X = XS_ To XE_
                RGBTriB = Dest24(Dest1D)
                Dest24(Dest1D).Blue = RGBTriB.Blue + m_alpha * (Src24(X).Blue - CLng(RGBTriB.Blue))
                Dest24(Dest1D).Green = RGBTriB.Green + m_alpha * (Src24(X).Green - CLng(RGBTriB.Green))
                Dest24(Dest1D).Red = RGBTriB.Red + m_alpha * (Src24(X).Red - CLng(RGBTriB.Red))
                Dest1D = Dest1D + 1
            Next
            SrcPos_ = X
        End If
    Next
End Sub

Private Sub m_MaskBlit_D24_S32(ByVal Dest1D&, ByVal SrcPos_&, StopRight_&, X&, XS_&, XE_&, ML_ As MaskInfo, Dest24() As RGBTriple, SrcAry() As Long)
Dim Seg_&
    For Seg_ = 0 To ML_.PartsM
        XS_ = ML_.DrawPart(Seg_).X1
        XE_ = ML_.DrawPart(Seg_).X2
        If XE_ >= SrcPos_ Then
            If XS_ < SrcPos_ Then
                XS_ = SrcPos_
            ElseIf XS_ > SrcPos_ Then
                Dest1D = Dest1D + XS_ - SrcPos_
                SrcPos_ = XS_
            End If
            If XE_ >= StopRight_ Then
                For X = XS_ To StopRight_
                    FGColor = SrcAry(X)
                    Dest24(Dest1D).Blue = FGColor And &HFF&
                    Dest24(Dest1D).Green = (FGColor And &HFF00&) / 256&
                    Dest24(Dest1D).Red = (FGColor And &HFF0000) / 65536
                    Dest1D = Dest1D + 1
                Next
                Exit Sub
            End If
            For X = XS_ To XE_
                FGColor = SrcAry(X)
                Dest24(Dest1D).Blue = FGColor And &HFF&
                Dest24(Dest1D).Green = (FGColor And &HFF00&) / 256&
                Dest24(Dest1D).Red = (FGColor And &HFF0000) / 65536
                Dest1D = Dest1D + 1
            Next
            SrcPos_ = X
        End If
    Next
End Sub
Private Sub m_MaskBlit_D24_S32_Alpha(ByVal Dest1D&, ByVal SrcPos_&, StopRight_&, X&, XS_&, XE_&, ML_ As MaskInfo, Dest24() As RGBTriple, SrcAry() As Long)
Dim Seg_&
    For Seg_ = 0 To ML_.PartsM
        XS_ = ML_.DrawPart(Seg_).X1
        XE_ = ML_.DrawPart(Seg_).X2
        If XE_ >= SrcPos_ Then
            If XS_ < SrcPos_ Then
                XS_ = SrcPos_
            ElseIf XS_ > SrcPos_ Then
                Dest1D = Dest1D + XS_ - SrcPos_
                SrcPos_ = XS_
            End If
            If XE_ >= StopRight_ Then
                For X = XS_ To StopRight_
                    RGBTriA = Dest24(Dest1D)
                    FGColor = SrcAry(X)
                    Dest24(Dest1D).Blue = RGBTriA.Blue + m_alpha * ((FGColor And &HFF&) - CLng(RGBTriA.Blue))
                    Dest24(Dest1D).Green = RGBTriA.Green + m_alpha * ((FGColor And &HFF00&) / 256& - CLng(RGBTriA.Green))
                    Dest24(Dest1D).Red = RGBTriA.Red + m_alpha * ((FGColor And &HFF0000) / 65536 - CLng(RGBTriA.Red))
                    Dest1D = Dest1D + 1
                Next
                Exit Sub
            End If
            For X = XS_ To XE_
                RGBTriA = Dest24(Dest1D)
                FGColor = SrcAry(X)
                Dest24(Dest1D).Blue = RGBTriA.Blue + m_alpha * ((FGColor And &HFF&) - CLng(RGBTriA.Blue))
                Dest24(Dest1D).Green = RGBTriA.Green + m_alpha * ((FGColor And &HFF00&) / 256& - CLng(RGBTriA.Green))
                Dest24(Dest1D).Red = RGBTriA.Red + m_alpha * ((FGColor And &HFF0000) / 65536 - CLng(RGBTriA.Red))
                Dest1D = Dest1D + 1
            Next
            SrcPos_ = X
        End If
    Next
End Sub

Private Sub m_MaskBlit_D32_S32_Alpha(ByVal Dest1D&, s_alpha!, ByVal SrcPos_&, StopRight_&, XS_&, XE_&, ML_ As MaskInfo, Dest1() As Long, Src1() As Long)
Dim Seg_&, X&

    For Seg_ = 0 To ML_.PartsM
        XS_ = ML_.DrawPart(Seg_).X1
        XE_ = ML_.DrawPart(Seg_).X2
        If XE_ >= SrcPos_ Then
            If XS_ < SrcPos_ Then
                XS_ = SrcPos_
            ElseIf XS_ > SrcPos_ Then
                Dest1D = Dest1D + XS_ - SrcPos_
                SrcPos_ = XS_
            End If
            If XE_ >= StopRight_ Then
                For X = XS_ To StopRight_
                    BGColor = Dest1(Dest1D) And &HFFFFFF
                    FGColor = Src1(X)
                    Dest1(Dest1D) = BGColor + _
                     (s_alpha * ((FGColor And &HFF0000) - (BGColor And &HFF0000)) And &HFF0000 Or _
                     s_alpha * ((FGColor And &HFF00&) - (BGColor And &HFF00&)) And &HFF00& Or _
                     s_alpha * ((FGColor And &HFF&) - (BGColor And &HFF&)) And &HFF&)
                    Dest1D = Dest1D + 1
                Next
                Exit Sub
            End If
            For X = XS_ To XE_
                BGColor = Dest1(Dest1D) And &HFFFFFF
                FGColor = Src1(X)
                Dest1(Dest1D) = BGColor + _
                 (s_alpha * ((FGColor And &HFF0000) - (BGColor And &HFF0000)) And &HFF0000 Or _
                 s_alpha * ((FGColor And &HFF00&) - (BGColor And &HFF00&)) And &HFF00& Or _
                 s_alpha * ((FGColor And &HFF&) - (BGColor And &HFF&)) And &HFF&)
                Dest1D = Dest1D + 1
            Next
            SrcPos_ = X
        End If
    Next
End Sub
Private Sub m_MaskBlit_D32_S24(ByVal Dest1D&, s_alpha!, ByVal SrcPos_&, StopRight_&, XS_&, XE_&, ML_ As MaskInfo, Dest1() As Long, Src1() As RGBTriple)
Dim Seg_&, X&

    For Seg_ = 0 To ML_.PartsM
        XS_ = ML_.DrawPart(Seg_).X1
        XE_ = ML_.DrawPart(Seg_).X2
        If XE_ >= SrcPos_ Then
            If XS_ < SrcPos_ Then
                XS_ = SrcPos_
            ElseIf XS_ > SrcPos_ Then
                Dest1D = Dest1D + XS_ - SrcPos_
                SrcPos_ = XS_
            End If
            If XE_ >= StopRight_ Then
                For X = XS_ To StopRight_
                    Dest1(Dest1D) = 65536 * Src1(X).Red Or _
                     256& * Src1(X).Green Or _
                     Src1(X).Blue
                    Dest1D = Dest1D + 1
                Next
                Exit Sub
            End If
            For X = XS_ To XE_
                Dest1(Dest1D) = 65536 * Src1(X).Red Or _
                 256& * Src1(X).Green Or _
                 Src1(X).Blue
                Dest1D = Dest1D + 1
            Next
            SrcPos_ = X
        End If
    Next
End Sub
Private Sub m_MaskBlit_D32_S24_Alpha(ByVal Dest1D&, s_alpha!, ByVal SrcPos_&, StopRight_&, XS_&, XE_&, ML_ As MaskInfo, Dest1() As Long, Src1() As RGBTriple)
Dim Seg_&, X&

    For Seg_ = 0 To ML_.PartsM
        XS_ = ML_.DrawPart(Seg_).X1
        XE_ = ML_.DrawPart(Seg_).X2
        If XE_ >= SrcPos_ Then
            If XS_ < SrcPos_ Then
                XS_ = SrcPos_
            ElseIf XS_ > SrcPos_ Then
                Dest1D = Dest1D + XS_ - SrcPos_
                SrcPos_ = XS_
            End If
            If XE_ >= StopRight_ Then
                For X = XS_ To StopRight_
                    BGColor = Dest1(Dest1D) And &HFFFFFF
                    Dest1(Dest1D) = BGColor + _
                     (s_alpha * (65536 * Src1(X).Red - (BGColor And &HFF0000)) And &HFF0000 Or _
                     s_alpha * (256& * Src1(X).Green - (BGColor And &HFF00&)) And &HFF00& Or _
                     s_alpha * (Src1(X).Blue - (BGColor And &HFF&)) And &HFF&)
                    Dest1D = Dest1D + 1
                Next
                Exit Sub
            End If
            For X = XS_ To XE_
                BGColor = Dest1(Dest1D) And &HFFFFFF
                Dest1(Dest1D) = BGColor + _
                 (s_alpha * (65536 * Src1(X).Red - (BGColor And &HFF0000)) And &HFF0000 Or _
                 s_alpha * (256& * Src1(X).Green - (BGColor And &HFF00&)) And &HFF00& Or _
                 s_alpha * (Src1(X).Blue - (BGColor And &HFF&)) And &HFF&)
                Dest1D = Dest1D + 1
            Next
            SrcPos_ = X
        End If
    Next
End Sub
Private Sub m_MaskBlitReverse(ByVal Dest1D&, ByVal SrcPos&, StopRight&, X&, XS&, XE&, ML As MaskInfo, Dest1 As SurfaceWrapper, Src1 As SurfaceWrapper)
Dim Seg&
    For Seg = 0 To ML.PartsM
        XS = ML.DrawPart(Seg).X1
        XE = ML.DrawPart(Seg).X2
        If XE >= SrcPos Then
            If XS < SrcPos Then
                XS = SrcPos
            ElseIf XS > SrcPos Then
                Dest1D = Dest1D + XS - SrcPos
                SrcPos = XS
            End If
            If XE >= StopRight Then
                For X = XS To StopRight
                    Src1.Dib32(X) = Dest1.Dib32(Dest1D)
                    Dest1D = Dest1D + 1
                Next
                Exit Sub
            End If
            For X = XS To XE
                Src1.Dib32(X) = Dest1.Dib32(Dest1D)
                Dest1D = Dest1D + 1
            Next
            SrcPos = X
        End If
    Next
End Sub
Private Sub m_MaskBlitSolid(ByVal Dest1D&, ByVal SrcPos_&, StopRight_&, X&, XS_&, XE_&, ML_ As MaskInfo, Dest1 As SurfaceWrapper, Src1 As SurfaceWrapper, Color&)
Dim Seg_&
    For Seg_ = 0 To ML_.PartsM
        XS_ = ML_.DrawPart(Seg_).X1
        XE_ = ML_.DrawPart(Seg_).X2
        If XE_ >= SrcPos_ Then
            If XS_ < SrcPos_ Then
                XS_ = SrcPos_
            ElseIf XS_ > SrcPos_ Then
                Dest1D = Dest1D + XS_ - SrcPos_
                SrcPos_ = XS_
            End If
            If XE_ >= StopRight_ Then
                For X = XS_ To StopRight_
                    Dest1.Dib32(Dest1D) = Color
                    Dest1D = Dest1D + 1
                Next
                Exit Sub
            End If
            For X = XS_ To XE_
                Dest1.Dib32(Dest1D) = Color
                Dest1D = Dest1D + 1
            Next
            SrcPos_ = X
        End If
    Next
End Sub
Private Function m_MakeSurf(Surf As SurfaceWrapper, ByVal Width%, ByVal Height%) As Boolean
Dim TmpTot&

 Width = Abs(Width)
 Height = Abs(Height)
 
 TmpTot = CLng(Width) * CLng(Height)
 
 m_MakeSurf = False
 
 If TmpTot > 0 Then
    Surf.BM.bmWidth = Width
    Surf.BM.bmHeight = Height
    ClearSurface Surf
    m_SetProcDims Surf
    ReDim Surf.Dib32(Surf.UB32)
    m_MakeSurf = True
 End If
 
End Function
Private Sub m_SetProcDims(Surf As SurfaceWrapper)

 Surf.WideM = Surf.BM.bmWidth - 1
 Surf.HighM = Surf.BM.bmHeight - 1
 
 Surf.TotalPixels = Surf.BM.bmWidth * Surf.BM.bmHeight
 
 Surf.Proc.BytesPixel = Surf.BM.bmBitsPixel / 8
 Surf.Proc.PixelTopLeft = Surf.BM.bmWidth * Surf.HighM
 Surf.Proc.BlueTopLeft = Surf.Proc.PixelTopLeft * Surf.Proc.BytesPixel
 Surf.Proc.BlueTopRight = Surf.Proc.BlueTopLeft + Surf.WideM * Surf.Proc.BytesPixel
 Surf.Proc.BlueRight = Surf.WideM * Surf.Proc.BytesPixel
 
 Surf.UB32 = Surf.TotalPixels - 1

With Surf.BMI.bmiHeader
.biSize = Len(Surf.BMI.bmiHeader)
.biWidth = Surf.BM.bmWidth
.biHeight = Surf.BM.bmHeight
.biPlanes = 1
.biBitCount = Surf.BM.bmBitsPixel
End With
 
End Sub
Private Sub m_SetMask(Surf As SurfaceWrapper, MaskInfo1 As MaskInfo, HPtr_&, BlitPixel As Boolean, ARGB1&, SrcX1&, Optional LY&)
Dim LX&
Dim RGBTRI As RGBTriple

    If Surf.BM.bmBitsPixel = 32 Then

        For SrcX1 = SrcX1 To SrcX1 + Surf.WideM
            If Surf.Dib32(SrcX1) = ARGB1 Then 'pixel does not get drawn
                If BlitPixel Then
                    MaskInfo1.DrawPart(HPtr_).X2 = SrcX1 - 1
                    BlitPixel = False
                End If
            Else 'this pixel is not masked (is supposed to be drawn)
                If Not BlitPixel Then
                    HPtr_ = HPtr_ + 1
                    MaskInfo1.DrawPart(HPtr_).X1 = SrcX1
                    BlitPixel = True
                End If
            End If
            LX = LX + 1
        Next
    
    ElseIf Surf.BM.bmBitsPixel = 24 Then
    
        RGBTRI.Blue = ARGB1 And &HFF&
        RGBTRI.Green = (ARGB1 And &HFF00&) / 256&
        RGBTRI.Red = (ARGB1 And &HFF0000) / 65536
        
        For SrcX1 = SrcX1 To SrcX1 + Surf.WideM
            m_SetMaskForScan24 HPtr_, BlitPixel, SrcX1, Surf.ScanLine24(LY).Row(SrcX1), RGBTRI, MaskInfo1
        Next
        
    End If
    
    If BlitPixel Then
        MaskInfo1.DrawPart(HPtr_).X2 = SrcX1 - 1
    End If
    
    MaskInfo1.PartsM = HPtr_

End Sub
Private Sub m_SetMaskForScan24(HPtr&, BlitPixel As Boolean, SrcX1&, ArrayTri As RGBTriple, RGBTRI As RGBTriple, MaskInfo1 As MaskInfo)

    If ArrayTri.Blue = RGBTRI.Blue And ArrayTri.Green = RGBTRI.Green And ArrayTri.Red = RGBTRI.Red Then
        If BlitPixel Then
            MaskInfo1.DrawPart(HPtr).X2 = SrcX1 - 1
            BlitPixel = False
        End If
    Else 'this pixel is not masked (is supposed to be drawn)
        If Not BlitPixel Then
            HPtr = HPtr + 1
            MaskInfo1.DrawPart(HPtr).X1 = SrcX1
            BlitPixel = True
        End If
    End If
    
End Sub
Private Sub m_SetScanLine24(SL_ As ScanLine, Ptr_&, BM As BITMAP, ScanLB&)
    SL_.SA.cbElements = 3
    SL_.SA.cDims = 1
    SL_.SA.pvData = Ptr_
    SL_.SA.cElements = BM.bmWidth
    SL_.SA.lLbound = ScanLB
    CopyMemory ByVal VarPtrArray(SL_.Row), VarPtr(SL_.SA), 4
End Sub
Private Sub m_CopyTo32(Surf As SurfaceWrapper, bBits() As Byte)
Dim BlueTopLeft24_&
    IA = 0
    BlueTopLeft24_ = Surf.BM.bmWidthBytes * Surf.HighM
    For DrawY = 0 To BlueTopLeft24_ Step Surf.BM.bmWidthBytes
        For DrawX = DrawY To DrawY + Surf.Proc.BlueRight Step Surf.BM.bmBitsPixel / 8
            Surf.Dib32(IA) = bBits(DrawX) Or bBits(DrawX + 1) * &H100& Or bBits(DrawX + 2) * &H10000
            IA = IA + 1
        Next
    Next
End Sub
Private Sub m_Adjust_strFileFolder(StrFolder1 As String)
    If StrFolder1 <> "" Then
     If Right$(StrFolder1, 1) <> "\" Then
      strFileFolder = StrFolder1 & "\"
     Else
      strFileFolder = StrFolder1
     End If
    End If
End Sub

Private Function m_GetRed(Color_&) As Long
    m_GetRed = (Color_ And &HFF0000) / 65536
End Function
Private Function m_GetGreen(Color_&) As Long
    m_GetGreen = (Color_ And &HFF00&) / 256
End Function
Private Function m_GetBlue(Color_&) As Long
    m_GetBlue = Color_ And &HFF&
End Function
Public Function ARGBHSV(hue_0_To_1530!, ByVal saturation_0_To_1!, value_0_To_255!) As Long
Dim hue_and_sat!
Dim value1!
Dim diff1!
Dim maxim!

 If value_0_To_255 > 0 Then
  value1 = value_0_To_255 + 0.5
  If saturation_0_To_1 > 0 Then
   maxim = hue_0_To_1530 - 1530& * Int(hue_0_To_1530 / 1530&)
   diff1 = saturation_0_To_1 * value_0_To_255
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    BGBlu = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     BGRed = Int(value1)
     BGGrn = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     BGGrn = Int(value1)
     BGRed = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    BGRed = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     BGGrn = Int(value1)
     BGBlu = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     BGBlu = Int(value1)
     BGGrn = Int(value1 - hue_and_sat)
    End If
   Else
    BGGrn = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     BGBlu = Int(value1)
     BGRed = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     BGRed = Int(value1)
     BGBlu = Int(value1 - hue_and_sat)
    End If
   End If
   ARGBHSV = BGRed * 65536 Or BGGrn * 256& Or BGBlu
  Else 'saturation_0_To_1 <= 0
   ARGBHSV = Int(value1) * GrayScaleRGB
  End If
 Else 'value_0_To_255 <= 0
  ARGBHSV = 0&
 End If
End Function


