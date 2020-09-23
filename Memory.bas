Attribute VB_Name = "modMemory"
Option Explicit

'Memory management module

'Written by Keith R. Weimer
'Way Too Happy Software

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Function GetAddress(ByVal Address As Long) As Long
    GetAddress = Address
End Function

Function Ceiling(ByVal Value As Double) As Long
    If Value - Int(Value) > 0 Then
        Ceiling = CLng(Value) + 1
    Else
        Ceiling = CLng(Value)
    End If
End Function

Function IsOdd(ByVal Value As Long) As Boolean
    IsOdd = Value And 1
End Function

Function IsEven(ByVal Value As Long) As Boolean
    IsEven = (Value And -2) = Value
End Function

Function FixHex(ByVal HexString As String, ByVal Bytes As Byte) As String
    If Bytes > Len(HexString) Then
        FixHex = String$(Bytes - Len(HexString), "0") & HexString
    Else
        FixHex = HexString
    End If
End Function

Function SplitEx(ByVal Text As String) As String()
    Dim Start As Integer
    Dim Char As String * 1
    Dim IsIn As Boolean
    Dim InQuote As Boolean
    Dim CurArg As Integer
    Dim Temp() As String
    
    CurArg = -1
    For Start = 1 To Len(Text)
        Char = Mid$(Text, Start, 1)
        If Char = """" Then
            InQuote = Not InQuote
            If InQuote Then
                CurArg = CurArg + 1
                ReDim Preserve Temp(0 To CurArg) As String
                IsIn = True
            Else
                IsIn = False
            End If
        ElseIf Char = " " And Not InQuote Then
            IsIn = False
        Else
            If Not IsIn Then
                CurArg = CurArg + 1
                ReDim Preserve Temp(0 To CurArg) As String
                IsIn = True
            End If
            Temp(CurArg) = Temp(CurArg) + Char
        End If
    Next Start
    
    SplitEx = Temp
End Function

Function IsNullString(Value As String) As Boolean
    IsNullString = StrPtr(Value) = 0
End Function

Function FixedLengthString(ByVal Length As Long, Optional Value As String, Optional ByVal Padding As Byte) As String
    If Length = Len(Value) Then
        FixedLengthString = Value
    ElseIf Length > Len(Value) Then
        FixedLengthString = Value & String$(Length - Len(Value), Padding)
    ElseIf Length < Len(Value) Then
        FixedLengthString = Left$(Value, Length)
    End If
    
    'FixedLengthString = StrConv(FixedLengthString, vbFromUnicode)
End Function

Function LTrimNull(Value As String) As String
    Dim Start As Long
    
    Start = InStrRev(Value, vbNullChar)
    If Start = 0 Then
        LTrimNull = Value
    Else
        LTrimNull = Mid$(Value, Start + 1)
    End If
End Function

Function RTrimNull(Value As String) As String
    Dim Start As Long
    
    Start = InStr(1, Value, vbNullChar)
    If Start = 0 Then
        RTrimNull = Value
    Else
        RTrimNull = Left$(Value, Start - 1)
    End If
End Function

Function SwapEndianWord(ByVal Value As Integer) As Integer
    Dim InputData(0 To 1) As Byte
    Dim OutputData(0 To 1) As Byte
    Dim Index As Byte
    
    CopyMemory ByVal VarPtr(InputData(0)), Value, 2
    
    For Index = 0 To 1
        OutputData(1 - Index) = InputData(Index)
    Next Index
    
    CopyMemory SwapEndianWord, ByVal VarPtr(OutputData(0)), 2
End Function

Function SwapEndianDWord(ByVal Value As Long) As Long
    Dim InputData(0 To 3) As Byte
    Dim OutputData(0 To 3) As Byte
    Dim Index As Byte
    
    CopyMemory ByVal VarPtr(InputData(0)), Value, 4
    
    For Index = 0 To 3
        OutputData(3 - Index) = InputData(Index)
    Next Index
    
    CopyMemory SwapEndianDWord, ByVal VarPtr(OutputData(0)), 4
End Function

Function Word(ByVal LoByte As Byte, Optional ByVal HiByte As Byte) As Integer
    CopyMemory ByVal VarPtr(Word), LoByte, 1
    CopyMemory ByVal VarPtr(Word) + 1, HiByte, 1
End Function

Function DWord(ByVal LoWord As Integer, Optional ByVal HiWord As Integer) As Long
    CopyMemory ByVal VarPtr(DWord), LoWord, 2
    CopyMemory ByVal VarPtr(DWord) + 2, HiWord, 2
End Function

Function LoByte(ByVal Word As Integer) As Byte
    CopyMemory LoByte, ByVal VarPtr(Word), 1
End Function

Function HiByte(ByVal Word As Integer) As Byte
    CopyMemory HiByte, ByVal VarPtr(Word) + 1, 1
End Function

Function LoWord(ByVal DWord As Long) As Integer
    CopyMemory LoWord, ByVal VarPtr(DWord), 2
End Function

Function HiWord(ByVal DWord As Long) As Integer
    CopyMemory HiWord, ByVal VarPtr(DWord) + 2, 2
End Function

Function WordToString(ByVal Value As Integer) As String
    Dim Data(0 To 1) As Byte
    
    CopyMemory ByVal VarPtr(Data(0)), Value, 2
    WordToString = StrConv(Data, vbUnicode)
End Function

Function StringToWord(ByVal Value As String) As Integer
    If Len(Value) = 2 Then
        Dim Data(0 To 1) As Byte
        Dim Index As Byte
        
        For Index = 0 To 1
            Data(Index) = Asc(Mid$(Value, Index + 1, 1))
        Next Index
        
        CopyMemory StringToWord, ByVal VarPtr(Data(0)), 2
    End If
End Function

Function DWordToString(ByVal Value As Long) As String
    Dim Data(0 To 3) As Byte
    
    CopyMemory ByVal VarPtr(Data(0)), Value, 4
    DWordToString = StrConv(Data, vbUnicode)
End Function

Function StringToDWord(ByVal Value As String) As Long
    If Len(Value) = 4 Then
        Dim Data(0 To 3) As Byte
        Dim Index As Byte
        
        For Index = 0 To 3
            Data(Index) = Asc(Mid$(Value, Index + 1, 1))
        Next Index
        
        CopyMemory StringToDWord, ByVal VarPtr(Data(0)), 4
    End If
End Function

Function ByteToBoolArray(ByVal InputData As Byte) As Boolean()
    Dim OutputData(0 To 7) As Boolean
    
    Dim InputBit As Byte
    
    For InputBit = 0 To 7
        OutputData(InputBit) = InputData And (2 ^ (7 - InputBit))
    Next InputBit
    
    ByteToBoolArray = OutputData
End Function

Function BoolArrayToByte(InputData() As Boolean) As Byte
    Dim OutputData(0 To 1) As Byte
    Dim OutputBit As Byte
    
    If LBound(InputData) = 0 And UBound(InputData) <= 7 Then
        For OutputBit = 0 To 7
            If OutputBit > UBound(InputData) Then
                Exit For
            Else
                If InputData(OutputBit) Then BoolArrayToByte = BoolArrayToByte Or (2 ^ (7 - OutputBit))
            End If
        Next OutputBit
    End If
End Function

Function WordToBoolArray(ByVal Word As Integer) As Boolean()
    Dim InputData(0 To 1) As Byte
    Dim OutputData(0 To 15) As Boolean
    
    Dim InputByte As Byte
    Dim InputBit As Byte
    
    CopyMemory ByVal VarPtr(InputData(0)), Word, 2
    
    For InputByte = 0 To 1
        For InputBit = 0 To 7
            OutputData(InputByte * 8 + InputBit) = InputData(InputByte) And (2 ^ (7 - InputBit))
        Next InputBit
    Next InputByte
    
    WordToBoolArray = OutputData
End Function

Function BoolArrayToWord(InputData() As Boolean) As Integer
    Dim OutputData(0 To 1) As Byte
    
    Dim OutputByte As Byte
    Dim OutputBit As Byte
    
    If LBound(InputData) = 0 And UBound(InputData) <= 15 Then
        For OutputByte = 0 To 1
            For OutputBit = 0 To 7
                If OutputByte * 8 + OutputBit > UBound(InputData) Then
                    Exit For
                Else
                    If InputData(OutputByte * 8 + OutputBit) Then OutputData(OutputByte) = OutputData(OutputByte) Or (2 ^ (7 - OutputBit))
                End If
            Next OutputBit
        Next OutputByte
    End If
    
    CopyMemory BoolArrayToWord, ByVal VarPtr(OutputData(0)), 2
End Function

Function DWordToBoolArray(ByVal DWord As Long) As Boolean()
    Dim InputData(0 To 3) As Byte
    Dim OutputData(0 To 31) As Boolean
    
    Dim InputByte As Byte
    Dim InputBit As Byte
    
    CopyMemory ByVal VarPtr(InputData(0)), DWord, 4
    
    For InputByte = 0 To 3
        For InputBit = 0 To 7
            OutputData(InputByte * 8 + InputBit) = InputData(InputByte) And (2 ^ (7 - InputBit))
        Next InputBit
    Next InputByte
    
    DWordToBoolArray = OutputData
End Function

Function BoolArrayToDWord(InputData() As Boolean) As Long
    Dim OutputData(0 To 3) As Byte
    
    Dim OutputByte As Byte
    Dim OutputBit As Byte
    
    If LBound(InputData) = 0 And UBound(InputData) <= 31 Then
        For OutputByte = 0 To 3
            For OutputBit = 0 To 7
                If OutputByte * 8 + OutputBit > UBound(InputData) Then
                    Exit For
                Else
                    If InputData(OutputByte * 8 + OutputBit) Then OutputData(OutputByte) = OutputData(OutputByte) Or (2 ^ (7 - OutputBit))
                End If
            Next OutputBit
        Next OutputByte
    End If
        
    CopyMemory BoolArrayToDWord, ByVal VarPtr(OutputData(0)), 4
End Function

Function StringToBoolArray(ByVal InputData As String) As Boolean()
    Dim OutputData() As Boolean
    ReDim OutputData(0 To Len(InputData) * 8 - 1)
    
    Dim InputByte As Byte
    Dim InputBit As Byte
    
    For InputByte = 0 To Len(InputData) - 1
        For InputBit = 0 To 7
            OutputData(InputByte * 8 + InputBit) = Asc(Mid$(InputData, InputByte + 1, 1)) And (2 ^ (7 - InputBit))
        Next InputBit
    Next InputByte
    
    StringToBoolArray = OutputData
End Function

Function GetBit(ByVal Value As Long, ByVal Bit As Byte) As Boolean
    GetBit = Value And (2 ^ Bit)
End Function

Function SetBit(ByVal Value As Long, ByVal Bit As Byte, ByVal Status As Boolean) As Boolean
    If Status Then
        SetBit = Value Or (2 ^ Bit)
    Else
        SetBit = Value And Not (2 ^ Bit)
    End If
End Function

Function ShiftBits(ByVal Value As Variant, ByVal Amount As Integer) As Variant
    Dim Length As Integer
    
    Select Case VarType(Value)
        Case vbByte: Length = 1
        Case vbInteger: Length = 2
        Case vbLong: Length = 4
        Case Else: Exit Function
    End Select
    
    If Amount > 0 Then 'Shift right
        ShiftBits = Value \ (2 ^ Amount)
    ElseIf Amount < 0 Then 'Shift left
        ShiftBits = (Value * (2 ^ -Amount)) And (2 ^ (8 * Length) - 1)
    End If
End Function

Function MergeBoolArray(InputData1() As Boolean, InputData2() As Boolean) As Boolean()
    Dim OutputData() As Boolean
    Dim Index As Long
    
    ReDim OutputData(0 To UBound(InputData1) + UBound(InputData2) + 1)
    
    For Index = 0 To UBound(InputData1)
        OutputData(Index) = InputData1(Index)
    Next Index
    
    For Index = 0 To UBound(InputData2)
        OutputData(Index + UBound(InputData1) + 1) = InputData2(Index)
    Next Index
    
    MergeBoolArray = OutputData
End Function

Function GetArrayElementCount(InputArray As Variant) As Long
    'On Error Resume Next
    
    If IsArray(InputArray) Then GetArrayElementCount = UBound(InputArray) - LBound(InputArray) + 1
End Function

Function ArrayElementExists(InputArray As Variant, ByVal Index As Long) As Boolean
    'On Error Resume Next
    
    If IsArray(InputArray) Then ArrayElementExists = Index >= LBound(InputArray) And Index <= UBound(InputArray)
End Function

Function ArrayToString(InputArray As Variant, Optional Separator As String, Optional Binary As Boolean) As String
    Dim ArrayIndex As Long
    
    If IsArray(InputArray) Then
        If GetArrayElementCount(InputArray) > 0 Then
            For ArrayIndex = LBound(InputArray) To UBound(InputArray)
                If Binary And VarType(InputArray(ArrayIndex)) = vbBoolean Then
                    ArrayToString = ArrayToString & IIf(InputArray(ArrayIndex), "1", "0") & Separator
                Else
                    ArrayToString = ArrayToString & InputArray(ArrayIndex) & Separator
                End If
            Next ArrayIndex
        End If
    End If
End Function

Function MidArray(InputArray As Variant, ByVal Start As Long, ByVal Length As Long) As Boolean()
    If IsArray(InputArray) Then
        Dim OutputArray() As Boolean
        ReDim OutputArray(0 To Length - 1)
        Dim Index As Long
        
        For Index = Start To Start + Length - 1 Step Sgn(Length)
            If Index > UBound(InputArray) Then
                Exit For
            Else
                OutputArray(Index - Start) = InputArray(Index)
            End If
        Next Index
        
        MidArray = OutputArray
    End If
End Function

Sub AddArrayItem(InputArray As Variant, Item As Variant, Optional ByVal Index As Long = -1)
    If IsArray(InputArray) Then
        If Index = -1 Then Index = GetArrayElementCount(InputArray)
        
        If GetArrayElementCount(InputArray) = 0 Then
            ReDim InputArray(0)
            InputArray(0) = Item
        Else
            Dim ArrayIndex As Long
            ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) + 1)
            
            For ArrayIndex = UBound(InputArray) To Index + 1 Step -1
                If IsObject(InputArray(ArrayIndex - 1)) Then
                    Set InputArray(ArrayIndex) = InputArray(ArrayIndex - 1)
                Else
                    InputArray(ArrayIndex) = InputArray(ArrayIndex - 1)
                End If
            Next ArrayIndex
            
            InputArray(Index) = Item
        End If
    End If
End Sub

Sub RemoveArrayItem(InputArray As Variant, ByVal Index As Long)
    If IsArray(InputArray) Then
        Select Case GetArrayElementCount(InputArray)
            Case 1: Erase InputArray
            Case Is > 1
                Dim ArrayIndex As Long
            
                For ArrayIndex = Index To UBound(InputArray) - 1
                    If IsObject(InputArray(ArrayIndex + 1)) Then
                        Set InputArray(ArrayIndex) = InputArray(ArrayIndex + 1)
                    Else
                        InputArray(ArrayIndex) = InputArray(ArrayIndex + 1)
                    End If
                Next ArrayIndex
                
                ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) - 1)
        End Select
    End If
End Sub
