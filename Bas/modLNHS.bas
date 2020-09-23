Attribute VB_Name = "modLNHS"
'##################################################################
'#                                                                #
'# LARGE NUMBER HANDLING SYSTEM (LNHS)                            #
'#                                                                #
'# A solution for handling large integers (up to 32768 digits)    #
'# in Visual Basic                                                #
'#                                                                #
'#----------------------------------------------------------------#
'#                                                                #
'# Copyright (C) 2002 by Philipp Emanuel Weidmann                 #
'#                                                                #
'##################################################################


' Code returned by functions when error occurs
Private Const ErrorCode = "ERROR"

Public Function LargeAdd(ByVal Number1 As String, ByVal Number2 As String) As String
On Error GoTo err
    ' Adds Number2 to Number1 and returns the result in a string
    Dim TempDigit1 As Integer, TempDigit2 As Integer
    Dim CalcResult As String
    Dim AddBuffer As Integer
    
    If Not IsNumeric(Number1) Or Not IsNumeric(Number2) Then
        LargeAdd = ErrorCode
        Exit Function
    End If
    
    ' Fill up the shorter number with zeroes at the beginning to make them equal in length
    If Len(Number1) > Len(Number2) Then
        Number2 = String(Len(Number1) - Len(Number2), "0") & Number2
    ElseIf Len(Number1) < Len(Number2) Then
        Number1 = String(Len(Number2) - Len(Number1), "0") & Number1
    End If
    
    ' Add one zero at the beginning of each number to make sure no digit gets lost when the
    ' sum of two digits is greater than 10 (see below)
    Number1 = "0" & Number1
    Number2 = "0" & Number2
    
    For cCount = Len(Number1) To 1 Step -1
        ' Add the numbers digit by digit to one another
        
        If Exponent1 Or Exponent2 Then Exit Function
        TempDigit1 = CInt(Mid(Number1, cCount, 1))
        TempDigit2 = CInt(Mid(Number2, cCount, 1))
        
        If TempDigit1 + TempDigit2 + AddBuffer >= 10 Then
            CalcResult = CStr(TempDigit1 + TempDigit2 + AddBuffer - 10) & CalcResult
            ' AddBuffer contains 1 if the sum of two digits is greater than 10
            AddBuffer = 1
        Else
            CalcResult = CStr(TempDigit1 + TempDigit2 + AddBuffer) & CalcResult
            AddBuffer = 0
        End If
    Next cCount
    
    If Left(CalcResult, 1) = "0" Then CalcResult = Mid(CalcResult, 2)
    LargeAdd = CalcResult
err:
    
End Function

Public Function LargeSubtract(ByVal Number1 As String, ByVal Number2 As String) As String
    ' Subtracts Number2 from Number1 and returns the result in a string
    Dim TempBuffer As String
    Dim TempDigit1 As Integer, TempDigit2 As Integer
    Dim CalcResult As String
    Dim SubtractBuffer As Integer
    
    If Not IsNumeric(Number1) Or Not IsNumeric(Number2) Then
        LargeSubtract = ErrorCode
        Exit Function
    End If
    
    ' If the numbers are equal, the result is zero
    If LargeCompare(Number1, Number2) = 0 Then
        LargeSubtract = "0"
        Exit Function
    End If
    
    ' Fill up the shorter number with zeroes at the beginning to make them equal in length
    If Len(Number1) > Len(Number2) Then
        Number2 = String(Len(Number1) - Len(Number2), "0") & Number2
    ElseIf Len(Number1) < Len(Number2) Then
        Number1 = String(Len(Number2) - Len(Number1), "0") & Number1
    End If
    
    ' Put the larger number in Number1 (otherwise digit-by-digit subtraction won't work)
    If LargeCompare(Number1, Number2) = 2 Then
        TempBuffer = Number2
        Number2 = Number1
        Number1 = TempBuffer
    End If
    
    For cCount = Len(Number1) To 1 Step -1
        ' Subtract the numbers digit by digit from one another
        TempDigit1 = CInt(Mid(Number1, cCount, 1))
        TempDigit2 = CInt(Mid(Number2, cCount, 1))
        
        If TempDigit1 - TempDigit2 - SubtractBuffer < 0 Then
            CalcResult = CStr(TempDigit1 - TempDigit2 - SubtractBuffer + 10) & CalcResult
            ' SubtractBuffer contains 1 if Digit1 - Digit2 < 0
            SubtractBuffer = 1
        Else
            CalcResult = CStr(TempDigit1 - TempDigit2 - SubtractBuffer) & CalcResult
            SubtractBuffer = 0
        End If
    Next cCount
    
    If TempBuffer <> "" Then CalcResult = "-" & CalcResult
    LargeSubtract = CalcResult
End Function

Public Function LargeMultiply(ByVal Number1 As String, ByVal Number2 As String) As String
    ' Multiplies Number1 with Number2 and returns the result in a string
    Dim TempDigit1 As Integer, TempDigit2 As Integer
    Dim CalcResult As String
    
    If Not IsNumeric(Number1) Or Not IsNumeric(Number2) Then
        LargeMultiply = ErrorCode
        Exit Function
    End If
    
    CalcResult = "0"
    
    For cCount = Len(Number1) To 1 Step -1
        For dCount = Len(Number2) To 1 Step -1
            TempDigit1 = CInt(Mid(Number1, cCount, 1))
            TempDigit2 = CInt(Mid(Number2, dCount, 1))
            
            ' Split the multiplication into additions:
            '
            ' abc * def = 10^0 * cf + 10^1 * ce + 10^2 * cd +
            '             10^1 * bf + 10^2 * be + 10^3 * bd +
            '             10^2 * af + 10^3 * ae + 10^4 * ad
            CalcResult = LargeAdd(CalcResult, CStr(TempDigit1 * TempDigit2) & _
                        String((Len(Number1) - cCount) + (Len(Number2) - dCount), "0"))
        Next dCount
    Next cCount
    
    LargeMultiply = CalcResult
End Function

Public Function LargeDivide(ByVal Number1 As String, ByVal Number2 As String) As String
    ' Divides Number1 by Number2 and returns the result in a string
    If Not IsNumeric(Number1) Or Not IsNumeric(Number2) Then
        LargeDivide = ErrorCode
        Exit Function
    End If
End Function

Public Function LargePower(ByVal Number1 As String, ByVal Number2 As String) As Integer
    ' Returns the result of Number1 ^ Number2 in a string
    Dim CalcResult As String
    
    If Not IsNumeric(Number1) Or Not IsNumeric(Number2) Then
        LargePower = ErrorCode
        Exit Function
    End If
    
    CalcResult = Number1
    
    ' I know this part is dirty and slow and could be done in much better ways,
    ' but I'm not a mathematician and at least it works ;-)
    Do While Not Number2 = "1"
        ' Decrement exponent by 1
        Number2 = LargeSubtract(Number2, "1")
        ' Multiply result with base
        CalcResult = LargeMultiply(CalcResult, Number1)
    Loop
    
    LargePower = CalcResult
End Function

Public Function LargeCompare(ByVal Number1 As String, ByVal Number2 As String) As Integer
    ' Compares Number1 with Number2
    '
    ' Returns 1 if Number1 is greater than Number2
    ' Returns 2 if Number2 is greater than Number1
    ' Returns 0 if the numbers are equal
    ' Returns -1 if an error occurs
    
    If Not IsNumeric(Number1) Or Not IsNumeric(Number2) Then
        LargeCompare = -1
        Exit Function
    End If
    
    ' Test if one of the numbers is shorter than the other one
    If Len(Number1) > Len(Number2) Then
        LargeCompare = 1
        Exit Function
    ElseIf Len(Number1) < Len(Number2) Then
        LargeCompare = 2
        Exit Function
    End If
    
    ' The numbers are equal in length => compare them digit by digit
    For cCount = 1 To Len(Number1)
        If CInt(Mid(Number1, cCount, 1)) > CInt(Mid(Number2, cCount, 1)) Then
            LargeCompare = 1
            Exit Function
        ElseIf CInt(Mid(Number1, cCount, 1)) < CInt(Mid(Number2, cCount, 1)) Then
            LargeCompare = 2
            Exit Function
        End If
    Next cCount
    
    LargeCompare = 0
End Function
