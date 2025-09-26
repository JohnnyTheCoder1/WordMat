Attribute VB_Name = "PlotCommandTests"
Option Explicit

' Basic tests for Plot Command functionality
' This module contains simple tests to validate plot command parsing and execution

Public Sub RunPlotTests()
    On Error GoTo TestError
    
    Dim testCount As Integer
    Dim passCount As Integer
    testCount = 0
    passCount = 0
    
    Debug.Print "=== Starting WordMat Plot Command Tests ==="
    
    ' Test 1: Check if basic plot command is recognized
    testCount = testCount + 1
    If TestLooksLikePlotCommand() Then
        passCount = passCount + 1
        Debug.Print "✓ Test 1 PASSED: LooksLikePlotCommand recognition"
    Else
        Debug.Print "✗ Test 1 FAILED: LooksLikePlotCommand recognition"
    End If
    
    ' Test 2: Parse basic plot command
    testCount = testCount + 1
    If TestBasicParsing() Then
        passCount = passCount + 1
        Debug.Print "✓ Test 2 PASSED: Basic plot command parsing"
    Else
        Debug.Print "✗ Test 2 FAILED: Basic plot command parsing"
    End If
    
    ' Test 3: Parse plot command with options
    testCount = testCount + 1
    If TestParsingWithOptions() Then
        passCount = passCount + 1
        Debug.Print "✓ Test 3 PASSED: Plot command parsing with options"
    Else
        Debug.Print "✗ Test 3 FAILED: Plot command parsing with options"
    End If
    
    ' Test 4: Expression validation
    testCount = testCount + 1
    If TestExpressionValidation() Then
        passCount = passCount + 1
        Debug.Print "✓ Test 4 PASSED: Expression validation"
    Else
        Debug.Print "✗ Test 4 FAILED: Expression validation"
    End If
    
    Debug.Print "=== Test Results: " & passCount & "/" & testCount & " tests passed ==="
    
    If passCount = testCount Then
        MsgBox "All plot command tests passed! (" & passCount & "/" & testCount & ")", vbInformation, "WordMat Plot Tests"
    Else
        MsgBox "Some tests failed. Passed: " & passCount & "/" & testCount & vbCrLf & "Check Debug window for details.", vbExclamation, "WordMat Plot Tests"
    End If
    
    Exit Sub
    
TestError:
    Debug.Print "✗ ERROR in RunPlotTests: " & Err.Description
    MsgBox "Error running tests: " & Err.Description, vbCritical, "WordMat Plot Tests"
End Sub

Private Function TestLooksLikePlotCommand() As Boolean
    On Error GoTo TestError
    
    TestLooksLikePlotCommand = False
    
    ' Test positive cases
    If Not LooksLikePlotCommand("plot(sin(x), -2*pi, 2*pi)") Then Exit Function
    If Not LooksLikePlotCommand("plot(x^2, -5, 5)") Then Exit Function
    If Not LooksLikePlotCommand("plot(cos(x), 0, 10; title=""Test"")") Then Exit Function
    
    ' Test negative cases
    If LooksLikePlotCommand("sin(x)") Then Exit Function
    If LooksLikePlotCommand("plot sin(x)") Then Exit Function
    If LooksLikePlotCommand("plot(sin(x), -2, 2") Then Exit Function ' Missing closing paren
    
    TestLooksLikePlotCommand = True
    Exit Function
    
TestError:
    Debug.Print "Error in TestLooksLikePlotCommand: " & Err.Description
    TestLooksLikePlotCommand = False
End Function

Private Function TestBasicParsing() As Boolean
    On Error GoTo TestError
    
    TestBasicParsing = False
    
    Dim cmd As PlotCommand
    
    ' Test basic command parsing
    If Not ParsePlotCommand("plot(sin(x), -2*pi, 2*pi)", cmd) Then Exit Function
    
    If cmd.Expression <> "sin(x)" Then Exit Function
    If cmd.XMin <> "-2*pi" Then Exit Function
    If cmd.XMax <> "2*pi" Then Exit Function
    If cmd.Title <> "" Then Exit Function
    
    TestBasicParsing = True
    Exit Function
    
TestError:
    Debug.Print "Error in TestBasicParsing: " & Err.Description
    TestBasicParsing = False
End Function

Private Function TestParsingWithOptions() As Boolean
    On Error GoTo TestError
    
    TestParsingWithOptions = False
    
    Dim cmd As PlotCommand
    
    ' Test command with options
    If Not ParsePlotCommand("plot(x^2, -5, 5; title=""Parabola"", grid=true, width=600, height=400, dpi=150)", cmd) Then Exit Function
    
    If cmd.Expression <> "x^2" Then Exit Function
    If cmd.XMin <> "-5" Then Exit Function
    If cmd.XMax <> "5" Then Exit Function
    If cmd.Title <> "Parabola" Then Exit Function
    If cmd.Grid <> True Then Exit Function
    If cmd.Width <> 600 Then Exit Function
    If cmd.Height <> 400 Then Exit Function
    If cmd.Dpi <> 150 Then Exit Function
    
    TestParsingWithOptions = True
    Exit Function
    
TestError:
    Debug.Print "Error in TestParsingWithOptions: " & Err.Description
    TestParsingWithOptions = False
End Function

Private Function TestExpressionValidation() As Boolean
    On Error GoTo TestError
    
    TestExpressionValidation = False
    
    ' Test valid expressions
    If Not IsValidExpression("sin(x)") Then Exit Function
    If Not IsValidExpression("x^2 + 2*x + 1") Then Exit Function
    If Not IsValidExpression("cos(3*x)") Then Exit Function
    If Not IsValidExpression("exp(-x^2)") Then Exit Function
    
    ' Test invalid expressions (these should fail validation)
    If IsValidExpression("system('rm -rf /')") Then Exit Function ' Security test
    If IsValidExpression("exec('malicious')") Then Exit Function ' Security test
    
    TestExpressionValidation = True
    Exit Function
    
TestError:
    Debug.Print "Error in TestExpressionValidation: " & Err.Description
    TestExpressionValidation = False
End Function

' Test function to validate plotting with a simple expression
Public Sub TestPlotExample()
    On Error GoTo TestError
    
    ' Create a simple test by inserting a plot command and executing it
    Selection.TypeText "plot(sin(x), -2*pi, 2*pi)"
    Selection.HomeKey unit:=wdLine, Extend:=wdExtend
    
    ' Try to execute the plot
    MsgBox "About to test plot execution. This will attempt to create a plot image.", vbInformation, "WordMat Plot Test"
    PlotSelection
    
    Exit Sub
    
TestError:
    MsgBox "Error in TestPlotExample: " & Err.Description, vbCritical, "WordMat Plot Test Error"
End Sub