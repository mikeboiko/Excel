'Fills all SCADA register addresses sequentially
Sub Fill_All_Addresses()

Application.ScreenUpdating = False
    
'Perform this macro on another workbook
Windows("PPL006-001 LAG P3 SCADA - IFC rev0.2.xlsx").Activate
    
'Initial Parameters
ShtAnalog = "Analog"
ShtRate = "Rate"
ShtStatus = "Status"

'Initial Register Values as assigned by SCADA
StartRangeAnalog = "J9"
StartRangeRate = "K18"
StartRangeStatusBits = "G12"
StartRangeStatusCommands = "K11"

RegValue = Fill_Analog(ShtAnalog, StartRangeAnalog)
Fill_Rate ShtRate, StartRangeRate, RegValue
Fill_Status_Bits ShtStatus, StartRangeStatusBits
Fill_Status_Commands ShtStatus, StartRangeStatusCommands

Application.ScreenUpdating = True

End Sub

'Fills all SCADA Analog addresses
Function Fill_Analog(Sht, StartRange)
    
    'Initialize
    Sheets(Sht).Select
    FirstRow = Range(StartRange).Row
    FirstCol = Range(StartRange).Column
    LastRow = Cells(100000, FirstCol).End(xlUp).Row
    RegValue = Replace(Cells(FirstRow, FirstCol).Value, " unsign 16 int", "") 'Initial Register Value
        
    'Loops through all addresses
    For Row = FirstRow + 1 To LastRow
        For Col = FirstCol To FirstCol + 3 Step 3 'The address columns are 3 apart
            'If string isn't blank and doesn't contain "PLC" (title), this is a register that needs to be renumbered
            If Cells(Row, Col).Value <> "" And InStr(Cells(Row, Col).Value, "PLC") = 0 Then
                RegValue = RegValue + 1
                Cells(Row, Col).Value = RegValue & " unsign 16 int"
            End If
            
        Next
    Next
    
Fill_Analog = RegValue

End Function

'Fills all SCADA Rate addresses
Function Fill_Rate(Sht, StartRange, RegVal)
    
    'Initialize
    Sheets(Sht).Select
    FirstRow = Range(StartRange).Row
    Col = Range(StartRange).Column
    LastRow = Cells(100000, Col).End(xlUp).Row
    RegValue = RegVal 'Passed from last register value in Analog spreadsheet
    
    'Offset required so the first register in Rate tab is 1 higher than the last one in Analog tab
    RegValue = RegValue - 1
    
    'Loops through all addresses
    For Row = FirstRow To LastRow
        'If string isn't blank and doesn't contain "PLC" (title), this is a register that needs to be renumbered
        If Cells(Row, Col).Value <> "" And InStr(Cells(Row, Col).Value, "PLC") = 0 Then
            RegValue = RegValue + 2
            Cells(Row, Col).Value = RegValue & " PDM"
        End If
    Next

End Function

'Fills all SCADA bit addresses
Function Fill_Status_Bits(Sht, StartRange)
    
    'Initialize
    Sheets(Sht).Select
    FirstRow = Range(StartRange).Row
    FirstCol = Range(StartRange).Column
    LastRow = Cells(100000, FirstCol).End(xlUp).Row
    RegValue = Cells(FirstRow, FirstCol).Value 'Initial Register Value
    
    'Extract SCADA Word and Bit from initial address
    AddressSplit = Split(RegValue, "/")
    SWord = AddressSplit(0)
    SBit = AddressSplit(1)
    
    'Loops through all addresses
    For Row = FirstRow + 1 To LastRow
        For Col = FirstCol To FirstCol + 1
            'If string isn't blank and doesn't contain "PLC" (title), this is a register that needs to be renumbered
            If Cells(Row, Col).Value <> "" And InStr(Cells(Row, Col).Value, "PLC") = 0 Then
                'Bits are incremented from 0 to 15 and then the next word goes down by 1
                IncrementArray = WordBitIncrement(SWord, SBit)
                RegValue = IncrementArray(0)
                SWord = IncrementArray(1)
                SBit = IncrementArray(2)
                Cells(Row, Col).Value = RegValue
            End If
            
        Next
    Next

End Function

'Fills all SCADA command register addresses
Function Fill_Status_Commands(Sht, StartRange)

    'Initialize
    Sheets(Sht).Select
    FirstRow = Range(StartRange).Row
    FirstCol = Range(StartRange).Column
    LastRow = Cells(100000, FirstCol).End(xlUp).Row
    RegValue = Cells(FirstRow, FirstCol).Value 'Initial Register Value
        
    'Loops through all addresses
    For Row = FirstRow + 1 To LastRow
        For Col = FirstCol To FirstCol + 1
            'If string isn't blank and doesn't contain "PLC" (title), this is a register that needs to be renumbered
            If Cells(Row, Col).Value <> "" And InStr(Cells(Row, Col).Value, "PLC") = 0 Then
                RegValue = RegValue + 1
                Cells(Row, Col).Value = RegValue
            End If
            
        Next
    Next

End Function

'Bits are incremented from 0 to 15 and then the next word goes down by 1
Function WordBitIncrement(SWord, SBit)

'Converts strings to numbers
LngWord = CLng(SWord)
IntBit = CInt(SBit)

'Increment bit by 1
IntBit = IntBit + 1

If IntBit = 16 Then
    LngWord = LngWord - 1
    IntBit = 0
End If

'Convert numbers to strings
StrWord = CStr(LngWord)
StrBit = CStr(IntBit)

'Add leading 0 for bits 0-9
If Len(StrBit) = 1 Then StrBit = "0" & StrBit

RegValue = StrWord & "/" & StrBit

Dim ReturnArray(3) As Variant
ReturnArray(0) = RegValue 'RegValue
ReturnArray(1) = StrWord 'SWord
ReturnArray(2) = StrBit 'SBit

WordBitIncrement = ReturnArray

End Function





'For Testing
Sub TestMacro()

FirstCol = 10
For Col = FirstCol To FirstCol + 3 Step 3
    Debug.Print Col
Next

End Sub

