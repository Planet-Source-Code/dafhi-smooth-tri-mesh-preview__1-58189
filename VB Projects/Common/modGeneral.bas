Attribute VB_Name = "modGeneral"
Option Explicit

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Dim IA&

Public Const pi As Double = 3.14159265358979
Public Const TwoPi As Double = 2 * pi
Public Const piBy2 As Single = pi / 2
Public Const halfPi As Single = piBy2

Public Const NOTE_1OF12 As Double = 2 ^ (1 / 12)
Public Const NOTE_2OF12 As Double = 2 ^ (2 / 12)
Public Const NOTE_3OF12 As Double = 2 ^ (3 / 12)
Public Const NOTE_4OF12 As Double = 2 ^ (4 / 12)
Public Const NOTE_5OF12 As Double = 2 ^ (5 / 12)
Public Const NOTE_6OF12 As Double = 2 ^ (6 / 12)
Public Const NOTE_7OF12 As Double = 2 ^ (7 / 12)
Public Const NOTE_8OF12 As Double = 2 ^ (8 / 12)
Public Const NOTE_9OF12 As Double = 2 ^ (9 / 12)
Public Const NOTE_10OF12 As Double = 2 ^ (10 / 12)
Public Const NOTE_11OF12 As Double = 2 ^ (11 / 12)

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Dim Tick&
Dim NextTick&

Dim LBA As Long
Dim UBA As Long
Dim LenA As Long

Private StrTemp As String

Declare Function timeGetTime Lib "winmm.dll" () As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd&, lprcUpdate As RECT, ByVal hrgnUpdate&, ByVal fuRedraw&) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Add(Varia1 As Variant, ByVal value_ As Double)
    Varia1 = Varia1 + value_
End Sub
Sub AngleModulus(ByRef retAngle As Single)
 retAngle = retAngle - TwoPi * Int(retAngle / TwoPi)
End Sub

Function CheckFPS(RetVal As Variant, FrameNum As Variant, Optional Interval_Millisec& = 1000) As Boolean
    FrameNum = FrameNum + 1
    CheckFPS = False
    Tick = timeGetTime
    If Tick > NextTick Then
        RetVal = FrameNum
        NextTick = Tick + Interval_Millisec - 1
        FrameNum = 0
        CheckFPS = True
    End If
End Function
Sub FillBytesFromString(Bytes1() As Byte, ByVal Str1 As String)
 LBA = LBound(Bytes1)
 UBA = UBound(Bytes1)
 StrTemp = Left$(Str1, UBA - LBA + 1)
 For IA = LBA To UBA
  Bytes1(IA) = Asc(Mid$(StrTemp, IA + 1, 1))
 Next
End Sub
Function StringFromBytes(Bytes() As Byte) As String
Dim J1&

 LenA = UBound(Bytes) - LBound(Bytes) + 1

 If LenA > 0 Then
  StringFromBytes = Bytes
  StringFromBytes = StringFromBytes + StringFromBytes
  J1 = 1
  For IA = LBound(Bytes) To UBound(Bytes)
   Mid$(StringFromBytes, J1, 1) = Chr$(Bytes(IA))
   J1 = J1 + 1
  Next IA
 End If
 
End Function
Function IsFile(strFileSpec As String) As Boolean
    If strFileSpec = "" Then Exit Function
    If Len(Dir$(strFileSpec)) > 0 Then
        IsFile = True
    Else
        IsFile = False
    End If
End Function
