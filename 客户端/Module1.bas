Attribute VB_Name = "Module1"
Option Explicit

Public zh, version, version_str, money, Debugmode, last_date, roomname, cards()
Public net_fa, net_shou


Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001




Function Utf8ToUnicode(ByRef Utf() As Byte) As String
    Dim lRet As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    lLength = UBound(Utf) - LBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2
    Utf8ToUnicode = String$(lBufferSize, Chr(0))
    lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(Utf8ToUnicode), lBufferSize)
    If lRet <> 0 Then
    Utf8ToUnicode = Left(Utf8ToUnicode, lRet)
    End If
End Function

Public Function UTF8_Encode(ByVal strUnicode As String) As Byte()
'UTF-8 ±àÂë

    Dim TLen As Long
    Dim lngBufferSize As Long
    Dim lngResult As Long
    Dim bytUtf8() As Byte
    
    TLen = Len(strUnicode)
    If TLen = 0 Then Exit Function
    
    lngBufferSize = TLen * 3 + 1
    ReDim bytUtf8(lngBufferSize - 1)
    
    lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strUnicode), TLen, bytUtf8(0), lngBufferSize, vbNullString, 0)
    
    If lngResult <> 0 Then
        lngResult = lngResult - 1
        ReDim Preserve bytUtf8(lngResult)
    End If
    
    UTF8_Encode = bytUtf8
End Function


