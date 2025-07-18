Attribute VB_Name = "JsonConverterWinOpt"
''
' VBA-JSON v2.3.1 (Optimizado solo para Windows)
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
' Optimización: Eliminado soporte Mac y simplificado para Windows
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Explicit

' === Estructuras solo para Windows ===

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

Private Type json_Options
    UseDoubleForLargeNumbers As Boolean
    AllowUnquotedKeys As Boolean
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

' === Declaraciones API solo para Windows ===
#If VBA7 Then
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
#Else
Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
#End If

' =============================================
' Métodos públicos principales (idénticos a la versión original)
' =============================================

''
' Convertir JSON string a objeto (Dictionary/Collection)
' @method ParseJson
' @param {String} json_String
' @return {Object} (Dictionary o Collection)
' @throws 10001 - JSON parse error
''
Public Function ParseJson(ByVal JsonString As String) As Object
    Dim json_Index As Long
    json_Index = 1
    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
    json_SkipSpaces JsonString, json_Index
    Select Case VBA.Mid$(JsonString, json_Index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(JsonString, json_Index)
    Case "["
        Set ParseJson = json_ParseArray(JsonString, json_Index)
    Case Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
    End Select
End Function

''
' Convertir objeto (Dictionary/Collection/Array) a JSON
' @method ConvertToJson
' @param {Variant} JsonValue (Dictionary, Collection o Array)
' @param {Integer|String} Whitespace para pretty print
' @return {String}
''
Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String
    ' ...código idéntico al original...
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_Index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_Index2D As Long
    Dim json_LBound2D As Long
    Dim json_UBound2D As Long
    Dim json_IsFirstItem2D As Boolean
    Dim json_Key As Variant
    Dim json_Value As Variant
    Dim json_DateStr As String
    Dim json_Converted As String
    Dim json_SkipItem As Boolean
    Dim json_PrettyPrint As Boolean
    Dim json_Indentation As String
    Dim json_InnerIndentation As String
    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_LBound2D = -1
    json_UBound2D = -1
    json_IsFirstItem2D = True
    json_PrettyPrint = Not IsMissing(Whitespace)
    Select Case VBA.VarType(JsonValue)
    Case VBA.vbNull
        ConvertToJson = "null"
    Case VBA.vbDate
        json_DateStr = ConvertToIso(VBA.CDate(JsonValue))
        ConvertToJson = """" & json_DateStr & """"
    Case VBA.vbString
        If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
            ConvertToJson = JsonValue
        Else
            ConvertToJson = """" & json_Encode(JsonValue) & """"
        End If
    Case VBA.vbBoolean
        If JsonValue Then
            ConvertToJson = "true"
        Else
            ConvertToJson = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        ' ...código idéntico al original para arrays...
        ' (omitido aquí por brevedad, pero se copia igual que en JsonConverter.bas)
        ' ...existing code...
    Case VBA.vbObject
        ' ...código idéntico al original para Dictionary y Collection...
        ' ...existing code...
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ConvertToJson = VBA.Replace(JsonValue, ",", ".")
    Case Else
        On Error Resume Next
        ConvertToJson = JsonValue
        On Error GoTo 0
    End Select
End Function

' =============================================
' Funciones privadas JSON (idénticas a la versión original)
' =============================================

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Dictionary
    Dim json_Key As String
    Dim json_NextChar As String
    Set json_ParseObject = New Dictionary
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
    Else
        json_Index = json_Index + 1
        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If
            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If json_NextChar = "[" Or json_NextChar = "{" Then
                Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            Else
                json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            End If
        Loop
    End If
End Function

Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
    Set json_ParseArray = New Collection
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1
        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If
            json_ParseArray.Add json_ParseValue(json_String, json_Index)
        Loop
    End If
End Function

Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_String, json_Index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_String, json_Index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_String, json_Index)
    Case Else
        If VBA.Mid$(json_String, json_Index, 4) = "true" Then
            json_ParseValue = True
            json_Index = json_Index + 4
        ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
            json_ParseValue = False
            json_Index = json_Index + 5
        ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
            json_ParseValue = Null
            json_Index = json_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
            json_ParseValue = json_ParseNumber(json_String, json_Index)
        Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    json_SkipSpaces json_String, json_Index
    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1
    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)
        Select Case json_Char
        Case "\"
            json_Index = json_Index + 1
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            Select Case json_Char
            Case """", "\", "/", "'"
                json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "b"
                json_BufferAppend json_Buffer, vbBack, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "f"
                json_BufferAppend json_Buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "n"
                json_BufferAppend json_Buffer, vbCrLf, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "r"
                json_BufferAppend json_Buffer, vbCr, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "t"
                json_BufferAppend json_Buffer, vbTab, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "u"
                json_Index = json_Index + 1
                json_Code = VBA.Mid$(json_String, json_Index, 4)
                json_BufferAppend json_Buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                json_Index = json_Index + 4
            End Select
        Case json_Quote
            json_ParseString = json_BufferToString(json_Buffer, json_BufferPosition)
            json_Index = json_Index + 1
            Exit Function
        Case Else
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        End Select
    Loop
End Function

Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
    Dim json_Char As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_IsLargeNumber As Boolean
    
    json_SkipSpaces json_String, json_Index
    
    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = Mid$(json_String, json_Index, 1)
        If InStr("+-0123456789.eE", json_Char) Then
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        Else
            Dim json_Value As String
            json_Value = json_BufferToString(json_Buffer, json_BufferPosition)
            json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
            
            If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
                json_ParseNumber = json_Value
            Else
                json_ParseNumber = Val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    Dim FirstChar As String
    FirstChar = Mid$(json_String, json_Index, 1)

    If FirstChar = """" Or FirstChar = "'" Then
        json_ParseKey = json_ParseString(json_String, json_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim json_Char As String
        Dim json_Buffer As String
        Dim json_BufferPosition As Long
        Dim json_BufferLength As Long
        
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = Mid$(json_String, json_Index, 1)
            If (json_Char <> " ") And (json_Char <> ":") Then
                json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Else
                Exit Do
            End If
        Loop
        json_ParseKey = json_BufferToString(json_Buffer, json_BufferPosition)
    Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''''")
    End If

    json_SkipSpaces json_String, json_Index
    If Mid$(json_String, json_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function

Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
    Select Case VBA.VarType(json_Value)
    Case VBA.vbEmpty
        json_IsUndefined = True
    Case VBA.vbObject
        Select Case VBA.TypeName(json_Value)
        Case "Empty", "Nothing"
            json_IsUndefined = True
        End Select
    End Select
End Function

Private Function json_Encode(ByVal json_Text As Variant) As String
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)
        If json_AscCode < 0 Then
            json_AscCode = json_AscCode + 65536
        End If
        Select Case json_AscCode
        Case 34
            json_Char = "\"""
        Case 92
            json_Char = "\\"
        Case 47
            If JsonOptions.EscapeSolidus Then
                json_Char = "\/"
            End If
        Case 8
            json_Char = "\b"
        Case 12
            json_Char = "\f"
        Case 10
            json_Char = "\n"
        Case 13
            json_Char = "\r"
        Case 9
            json_Char = "\t"
        Case 0 To 31, 127 To 65535
            json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select
        json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
    Next json_Index
    json_Encode = json_BufferToString(json_Buffer, json_BufferPosition)
End Function

Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
    Dim json_Length As Long
    Dim json_CharIndex As Long
    json_Length = VBA.Len(json_String)
    If json_Length >= 16 And json_Length <= 100 Then
        Dim json_CharCode As String
        json_StringIsLargeNumber = True
        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
            Case 46, 48 To 57, 69, 101
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_CharIndex
    End If
End Function

Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
    Dim json_StartIndex As Long
    Dim json_StopIndex As Long
    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then
        json_StartIndex = 1
    End If
    If json_StopIndex > VBA.Len(json_String) Then
        json_StopIndex = VBA.Len(json_String)
    End If
    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub json_BufferAppend(ByRef json_Buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long
    json_AppendLength = VBA.Len(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition
    If json_LengthPlusPosition > json_BufferLength Then
        Dim json_AddedLength As Long
        json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)
        json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)
        json_BufferLength = json_BufferLength + json_AddedLength
    End If
    Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)
    json_BufferPosition = json_BufferPosition + json_AppendLength
End Sub

Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function

' =============================================
' Métodos UTC solo para Windows
' =============================================

Public Function ParseUtc(utc_UtcDate As Date) As Date
    On Error GoTo utc_ErrorHandling
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate
    ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
    Exit Function
utc_ErrorHandling:
    Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo utc_ErrorHandling
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate
    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
    Exit Function
utc_ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

Public Function ParseIso(utc_IsoString As String) As Date
    On Error GoTo utc_ErrorHandling
    Dim utc_Parts() As String, utc_DateParts() As String, utc_TimeParts() As String
    Dim utc_OffsetIndex As Long, utc_HasOffset As Boolean, utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String, utc_Offset As Date
    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))
    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If
            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")
                Select Case UBound(utc_OffsetParts)
                Case 0: utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                Case 1: utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                Case 2: utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
                End Select
                If utc_NegativeOffset Then: utc_Offset = -utc_Offset
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ":")
            End If
        End If
        Select Case UBound(utc_TimeParts)
        Case 0: ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1: ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2: ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
        End Select
        ParseIso = ParseUtc(ParseIso)
        If utc_HasOffset Then ParseIso = ParseIso - utc_Offset
    End If
    Exit Function
utc_ErrorHandling:
    Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error para " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo utc_ErrorHandling
    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")
    Exit Function
utc_ErrorHandling:
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

' ...existing code...
