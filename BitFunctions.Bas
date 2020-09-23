Public Sub BitClearByte(bytNumber As Byte, bytBitNumber As Byte)
  ' Comments  : This procedure clears a bit in a byte variable
  ' Parameters: bytNumber - The byte variable to clear the bit in. The result
  '             is also returned in this parameter
  '             bytBitNumber - The bit number to clear
  ' Returns   : Nothing
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  bytNumber = bytNumber And Not Shlb(1, bytBitNumber)

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "BitClearByte"
  Resume PROC_EXIT

End Sub

Public Sub BitClearInteger(intNumber As Integer, bytBitNumber As Byte)
  ' Comments  : This procedure clears a bit in a long variable
  ' Parameters: intNumber - The integer variable to clear the bit in. The
  '                         result is also returned in this parameter
  '             bytBitNumber - The bit number to clear
  ' Returns   : Nothing
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  intNumber = intNumber And Not Shli(1, bytBitNumber)

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "BitClearInteger"
  Resume PROC_EXIT

End Sub

Public Sub BitClearLong(lngNumber As Long, bytBitNumber As Byte)
  ' Comments  : This procedure clears a bit in a long variable
  ' Parameters: lngNumber - The long variable to clear the bit in. The result
  '             is also returned in this parameter
  '             bytBitNumber - The bit number to clear
  ' Returns   : Nothing
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  lngNumber = lngNumber And Not Shll(1, bytBitNumber)

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, , _
    "BitClearLong"
  Resume PROC_EXIT

End Sub

Public Sub BitSetByte(bytNumber As Byte, bytBitNumber As Byte)
  ' Comments  : This procedure sets a bit in a byte variable
  ' Parameters: bytNumber - The byte variable to set the bit in. The result
  '             is also returned in this parameter
  '             bytBitNumber - The bit number to clear
  ' Returns   : Nothing
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  bytNumber = bytNumber Or Shlb(1, bytBitNumber)

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, , _
    "BitSetByte"
  Resume PROC_EXIT

End Sub

Public Sub BitSetInteger(intNumber As Integer, bytBitNumber As Byte)
  ' Comments  : This procedure sets a bit in an integer variable
  ' Parameters: intNumber - The integer variable to set the bit in. The result
  '             is also returned in this parameter
  '             bytBitNumber - The bit number to clear
  ' Returns   : Nothing
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  intNumber = intNumber Or Shli(1, bytBitNumber)

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, , _
    "BitSetInteger"
  Resume PROC_EXIT

End Sub

Public Sub BitSetLong(lngNumber As Long, bytBitNumber As Byte)
  ' Comments  : This procedure sets a bit in a long variable
  ' Parameters: lngNumber - The long variable to set the bit in. The result
  '             is also returned in this parameter
  '             bytBitNumber - The bit number to clear
  ' Returns   : Nothing
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  lngNumber = lngNumber Or Shll(1, bytBitNumber)

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, , _
    "BitSetLong"
  Resume PROC_EXIT

End Sub

Public Function BitTestByte(bytNumber As Byte, bytBitNumber As Byte) As Boolean
  ' Comments  : This procedure tests a bit to see if it is turned on
  ' Parameters: bytNumber - The byte variable to test the bit in
  '             bytBitNumber - The bit number to test
  ' Returns   : True if the bit is set, and False if it is not
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  BitTestByte = ((bytNumber And Shlb(1, bytBitNumber)) <> 0)

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, , _
    "BitTestByte"
  Resume PROC_EXIT

End Function

Public Function BitTestInteger( _
  intNumber As Integer, _
  bytBitNumber As Byte) _
  As Boolean
  ' Comments  : This procedure tests a bit to see if it is turned on
  ' Parameters: intNumber - The integer variable to test the bit in
  '             bytBitNumber - The bit number to test
  ' Returns   : True if the bit is set, and False if it is not
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  BitTestInteger = ((intNumber And Shli(1, bytBitNumber)) <> 0)

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "BitTestInteger"
  Resume PROC_EXIT

End Function

Public Function BitTestLong(lngNumber As Long, bytBitNumber As Byte) As Boolean
  ' Comments  : This procedure tests a bit to see if it is turned on
  ' Parameters: lngNumber - The long variable to test the bit in
  '             bytBitNumber - The bit number to test
  ' Returns   : True if the bit is set, and False if it is not
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  BitTestLong = ((lngNumber And Shll(1, bytBitNumber)) <> 0)

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, , _
    "BitTestLong"
  Resume PROC_EXIT

End Function

Public Function dblToLong(ByVal dblNumber As Double) As Long
  ' Comments  : This routine does an unsigned conversion from a double Value
  '             to a long Value. This procedure correctly handles any double
  '             value
  ' Parameters: dblNumber - the double value to convert to a long
  ' Returns   : long
  ' Source    : Total Visual SourceBook 2000
  '
  Dim dblDivisor As Double
  Dim dblTemp As Double

  On Error GoTo PROC_ERR

  ' Visual basic does not allow you enter the value &H100000000 directly,
  ' so we enter &H7FFFFFFF, double it and add two to create it.
  dblDivisor = &H7FFFFFFF
  dblDivisor = (dblDivisor * 2) + 2

  'if the number is larger than a long can store, then truncate it
  If dblNumber > dblDivisor Or dblNumber < 0 Then
    dblTemp = dblNumber - (Int(dblNumber / dblDivisor) * dblDivisor)
  Else
    dblTemp = dblNumber
  End If

  ' if the number is greater than a signed long, convert it to a
  ' negative
  If dblTemp > &H7FFFFFFF Then
    dblToLong = dblTemp - dblDivisor
  ElseIf dblTemp < 0 Then
    ' If the number is negative
    dblToLong = dblDivisor + dblTemp
  Else
    dblToLong = dblTemp
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "dblToLong"
  Resume PROC_EXIT

End Function

Public Function HiByte(ByVal intNumber As Integer) As Byte
  ' Comments  : Returns the high byte of the passed integer
  ' Parameters: intNumber - integer value for which to to return the high byte
  ' Returns   : byte
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  HiByte = Int((IntToLong(intNumber) / &H100&)) And &HFF&

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "HiByte"
  Resume PROC_EXIT

End Function

Public Function HiLong(ByVal dblNumber As Double) As Long
  ' Comments  : Returns the high long of the passed double
  ' Parameters: dblNumber - double value for which to return the high long
  ' Returns   : long
  ' Source    : Total Visual SourceBook 2000
  '
  Dim dblDivisor As Double

  On Error GoTo PROC_ERR

  ' Visual basic does not allow you enter the value &H100000000 directly,
  ' so we enter &H7FFFFFFF, double it and add two to create it.
  dblDivisor = &H7FFFFFFF
  dblDivisor = (dblDivisor * 2) + 2

  HiLong = dblToLong(Int((dblNumber / dblDivisor)))

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "HiLong"
  Resume PROC_EXIT

End Function

Public Function HiWord(lngNumber As Long) As Integer
  ' Comments  : Returns the high integer of the passed long
  ' Parameters: lngNumber - long value for which to return the high integer
  ' Returns   : integer
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  HiWord = LongToInt(Int((lngNumber / &H10000)))

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "HiWord"
  Resume PROC_EXIT

End Function

Public Function IntToByte(ByVal intNumber As Integer) As Byte
  ' Comments  : This routine does an unsigned conversion from an integer value
  '             to a byte value. This procedure correctly handles any integer
  '             value
  ' Parameters: intNumber - the integer Value to convert to a byte
  ' Returns   : Byte
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  IntToByte = intNumber And &HFF&

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "IntToByte"
  Resume PROC_EXIT

End Function

Public Function IntToLong(ByVal intNumber As Integer) As Long
  ' Comments  : This procedure converts an integer Value to a long Value,
  '             treating the integer as unsigned
  ' Parameters: intNumber - the integer to convert to long
  ' Returns   : long
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  ' This routine converts an integer Value to a long Value
  If intNumber < 0 Then
    IntToLong = intNumber + &H10000
  Else
    IntToLong = intNumber
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "IntToLong"
  Resume PROC_EXIT

End Function

Public Function LoByte(ByVal intNumber As Integer) As Byte
  ' Comments  : Returns the low byte of the passed integer
  ' Parameters: intNumber - integer value for which to return the low byte
  ' Returns   : byte
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  LoByte = intNumber And &HFF&

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LoByte"
  Resume PROC_EXIT

End Function

Public Function LoLong(ByVal dblNumber As Double) As Long
  ' Comments  : Returns the low long of the passed double
  ' Parameters: dblNumber - double value for which to return the low long
  ' Returns   : long
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  LoLong = dblToLong(dblNumber)

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LoLong"
  Resume PROC_EXIT

End Function

Public Function LongToDbl(ByVal lngNumber As Long) As Double
  ' Comments  : This routine converts a long Value to a double Value,
  '             treating the long as unsigned
  ' Parameters: lngNumber - the long to convert to double
  ' Returns   : double
  ' Source    : Total Visual SourceBook 2000
  '
  Dim dblDivisor As Double

  On Error GoTo PROC_ERR

  ' Visual basic does not allow you enter the value &H100000000 directly,
  ' so we enter &H7FFFFFFF, double it and add two to create it.
  dblDivisor = &H7FFFFFFF
  dblDivisor = (dblDivisor * 2) + 2

  If lngNumber < 0 Then
    LongToDbl = lngNumber + dblDivisor
  Else
    LongToDbl = lngNumber
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "LongToDbl"
  Resume PROC_EXIT

End Function

Public Function LongToInt(ByVal lngNumber As Long) As Integer
  ' Comments  : This routine does an unsigned conversion from a long Value
  '             to an integer value. This procedure correctly handles any
  '             long Value
  ' Parameters: lngNumber - the long value to convert to an integer
  ' Returns   : Integer
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  ' This routine converts a long Value to an integer
  lngNumber = lngNumber And &HFFFF&
  If lngNumber > &H7FFF Then
    LongToInt = lngNumber - &H10000
  Else
    LongToInt = lngNumber
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "LongToInt"
  Resume PROC_EXIT

End Function

Public Function LoWord(ByVal lngNumber As Long) As Integer
  ' Comments  : Returns the low integer of the passed long
  ' Parameters: lngNumber - long value for which to return the low integer
  ' Returns   : integer
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  LoWord = LongToInt(lngNumber And &HFFFF&)

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LoWord"
  Resume PROC_EXIT

End Function

Public Function MakeDbl( _
  ByVal lngHiLong As Long, _
  ByVal lngLoLong As Long) _
  As Double
  ' Comments  : Combines two longs into a double. This routine correctly
  '             handles negative input Values
  ' Parameters: lngHiLong - the high long
  '             lngLoLong - the low long
  ' Returns   : long
  ' Source    : Total Visual SourceBook 2000
  '
  Dim dblDivisor As Double

  On Error GoTo PROC_ERR

  ' Visual basic does not allow you enter the value &H100000000 directly,
  ' so we enter &H7FFFFFFF, double it and add two to create it.
  dblDivisor = &H7FFFFFFF
  dblDivisor = (dblDivisor * 2) + 2

  MakeDbl = Int(CDbl(LongToDbl(lngHiLong) * dblDivisor) + _
    LongToDbl(lngLoLong))

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "MakeDbl"
  Resume PROC_EXIT

End Function

Public Function MakeInt( _
  ByVal bytHiByte As Byte, _
  ByVal bytLoByte As Byte) _
  As Integer
  ' Comments  : Combines two bytes into an integer.
  ' Parameters: bytHiByte - the high byte
  '             bytLoByte - the low byte
  ' Returns   : integer
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  MakeInt = bytLoByte Or LongToInt((CLng(bytHiByte) * &H100))

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "MakeInt"
  Resume PROC_EXIT

End Function

Public Function MakeLong( _
  ByVal intHiWord As Integer, _
  ByVal intLoWord As Integer) _
  As Long
  ' Comments  : Combines two words (integers) into a long. This routine
  '             correctly handles negative input Values
  ' Parameters: intHiWord - the high integer
  '             intLoWord - the low integer
  ' Returns   : long
  ' Source    : Total Visual SourceBook 2000
  '
  On Error GoTo PROC_ERR

  MakeLong = (dblToLong(CDbl(IntToLong(intHiWord)) * &H10000)) + _
    IntToLong(intLoWord)

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "MakeLong"
  Resume PROC_EXIT

End Function

Public Function Shlb(ByVal bytValue As Byte, ByVal bytPlaces As Byte) As Byte
  ' Comments  : Shifts a numeric value left the specified number of bits.
  ' Parameters: bytValue - byte value to shift
  '             bytPlaces - number of places to shift
  ' Returns   : Shifted Value
  ' Source    : Total Visual SourceBook 2000
  '
  Dim lngMultiplier As Long

  On Error GoTo PROC_ERR

  ' if we are shifting 8 or more bits, then the result is always zero
  If bytPlaces >= 8 Then
    Shlb = 0
  Else
    lngMultiplier = 2 ^ bytPlaces
    Shlb = IntToByte(LongToInt(bytValue * lngMultiplier))
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "Shlb"
  Resume PROC_EXIT

End Function

Public Function Shli( _
  ByVal intValue As Integer, _
  ByVal bytPlaces As Byte) _
  As Integer
  ' Comments  : Shifts a numeric value left the specified number of bits.
  ' Parameters: intValue - integer value to shift
  '             bytPlaces - number of places to shift
  ' Returns   : Shifted Value
  ' Source    : Total Visual SourceBook 2000
  '
  Dim lngMultiplier As Long

  On Error GoTo PROC_ERR

  ' if we are shifting 16 or more bits, then the result is always zero
  If bytPlaces >= 16 Then
    Shli = 0
  Else
    lngMultiplier = 2 ^ bytPlaces
    Shli = LongToInt(intValue * lngMultiplier)
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "Shli"
  Resume PROC_EXIT

End Function

Public Function Shll(ByVal lngNumber As Long, ByVal bytPlaces As Byte) As Long
  ' Comments  : Shifts a numeric Value left the specified number of bits.
  ' Parameters: lngNumber - long Value to shift
  '             bytPlaces - number of places to shift
  ' Returns   : Shifted Value
  ' Source    : Total Visual SourceBook 2000
  '
  Dim dblMultiplier As Double

  On Error GoTo PROC_ERR

  ' if we are shifting 32 or more bits, then the result is always zero
  If bytPlaces >= 32 Then
    Shll = 0
  Else
    dblMultiplier = 2 ^ bytPlaces
    Shll = dblToLong(lngNumber * dblMultiplier)
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "Shll"
  Resume PROC_EXIT

End Function

Public Function Shrb(ByVal bytValue As Byte, ByVal bytPlaces As Byte) As Byte
  ' Comments  : Shifts a byte value right the selected number of places
  '
  ' Parameters: bytValue - byte value to shift
  '             bytPlaces - number of places to shift
  ' Returns   : Shifted Value
  ' Source    : Total Visual SourceBook 2000
  '
  Dim lngDivisor As Long

  On Error GoTo PROC_ERR

  ' if we are shifting 8 or more bits, then the result is always zero
  If bytPlaces >= 8 Then
    Shrb = 0
  Else
    lngDivisor = 2 ^ bytPlaces
    Shrb = Int(bytValue / lngDivisor)
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "Shrb"
  Resume PROC_EXIT

End Function

Public Function Shri( _
  ByVal lngValue As Long, _
  ByVal bytPlaces As Byte) _
  As Integer
  ' Comments  : Shifts a long Value right the selected number of places
  ' Parameters: lngValue - integer Value to shift
  '             bytPlaces - number of places to shift
  ' Returns   : Shifted value
  ' Source    : Total Visual SourceBook 2000
  '
  Dim lngDivisor As Long

  On Error GoTo PROC_ERR

  ' if we are shifting 16 or more bits, then the result is always zero
  If bytPlaces >= 16 Then
    Shri = 0
  Else
    lngDivisor = 2 ^ bytPlaces
    Shri = Int(IntToLong(lngValue) / lngDivisor)
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Shri"
  Resume PROC_EXIT

End Function

Public Function Shrl(ByVal lngNumber As Long, ByVal bytPlaces As Byte) As Long
  ' Comments  : Shifts a long Value right the selected number of places
  ' Parameters: lngNumber - long Value to shift
  '             bytPlaces - number of places to shift
  ' Returns   : Shifted Value
  ' Source    : Total Visual SourceBook 2000
  '
  Dim lngDivisor As Long

  On Error GoTo PROC_ERR

  ' if we are shifting 32 or more bits, then the result is always zero
  If bytPlaces >= 32 Then
    Shrl = 0
  Else
    lngDivisor = 2 ^ bytPlaces
    Shrl = Int(LongToDbl(lngNumber) / lngDivisor)
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.number & ". " & Err.Description, ,  _
    "Shrl"
  Resume PROC_EXIT

End Function
