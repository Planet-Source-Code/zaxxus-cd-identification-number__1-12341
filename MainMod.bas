Attribute VB_Name = "Mainmod"
' Let's Declare Some Things...

Type toc
    min As Long
    sec As Long
    fram As Long
    offset As Long
End Type

Global cdtoc() As toc
Global totaltr As Integer
Global r As String * 40
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

' This Code Shows You Error Messages. (If There Are Any)

Public Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean

Static rc As Long
Static errStr As String * 200
rc = mciSendString(cmd, 0, 0, hWnd)
If (fShowError And rc <> 0) Then
mciGetErrorString rc, errStr, Len(errStr)
MsgBox errStr
End If
SendMCIString = (rc = 0)

End Function

' This Code Reads the CD's Table Of Contents (TOC).

Public Function readcdtoc() As Integer

mciSendString "status cd69 number of tracks wait", r, Len(r), 0
totaltr = CInt(Mid$(r, 1, 2))
ReDim cdtoc(totaltr + 1) As toc
mciSendString "set cd69 time format msf", 0, 0, 0
For i = 1 To totaltr
cmd = "status cd69 position track " & i
mciSendString cmd, r, Len(r), 0
cdtoc(i - 1).min = CInt(Mid$(r, 1, 2))
cdtoc(i - 1).sec = CInt(Mid$(r, 4, 2))
cdtoc(i - 1).fram = CInt(Mid$(r, 7, 2))
cdtoc(i - 1).offset = (cdtoc(i - 1).min * 60 * 75) + (cdtoc(i - 1).sec * 75) + cdtoc(i - 1).fram
Next

End Function

Public Function cddbsum(n) As Integer

ret = 0
m = n
For i = 1 To m
ret = ret + (n Mod 10)
n = n / 10
Next
cddbsum = ret

End Function

'This Code Gets The CD ID.

Public Function cddbdiscid(tr) As String

Dim n As Long
Dim tm As Long
For i = 0 To tr - 1
tm = ((cdtoc(i).min * 60) + cdtoc(i).sec)
Do While tm > 0
n = n + (tm Mod 10)
tm = tm \ 10
Loop
Next
mciSendString "status cd69 length wait", r, Len(r), 0
t = (CInt(Mid$(r, 1, 2)) * 60) + CInt(Mid$(r, 4, 2))
cddbdiscid = LCase$(Zeros(Hex$(n Mod &HFF), 2) & Zeros(Hex$(t), 4) & Zeros(Hex$(tr), 2))

End Function

Private Function Zeros(s As String, n As Integer) As String

If Len(s) < n Then
Zeros = String$(n - Len(s), "0") & s
Else
Zeros = s
End If

End Function
