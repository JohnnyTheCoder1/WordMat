Attribute VB_Name = "PlotDemo"
Option Explicit

' Demo module to showcase the plot(...) command functionality
' This creates sample plot commands for testing and demonstration

Public Sub CreatePlotDemoDocument()
    On Error GoTo DemoError
    
    ' Clear the document and add demonstration content
    ActiveDocument.Content.Delete
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Add title
    Selection.Font.Size = 16
    Selection.Font.Bold = True
    Selection.TypeText "WordMat Plot Command Demonstration" & vbCrLf & vbCrLf
    
    ' Reset formatting
    Selection.Font.Size = 12
    Selection.Font.Bold = False
    
    ' Add introduction
    Selection.TypeText "This document demonstrates the new plot(...) command feature in WordMat." & vbCrLf
    Selection.TypeText "To use: Select any plot command below and click 'Plot Selection' or use keyboard shortcut." & vbCrLf & vbCrLf
    
    ' Example 1: Basic sine function
    Selection.Font.Bold = True
    Selection.TypeText "Example 1: Basic Sine Function" & vbCrLf
    Selection.Font.Bold = False
    Selection.TypeText "plot(sin(x), -2*pi, 2*pi)" & vbCrLf & vbCrLf
    
    ' Example 2: Parabola with options
    Selection.Font.Bold = True
    Selection.TypeText "Example 2: Parabola with Options" & vbCrLf
    Selection.Font.Bold = False
    Selection.TypeText "plot(x^2, -5, 5; title=""Parabola"", grid=true, width=600, height=400)" & vbCrLf & vbCrLf
    
    ' Example 3: Cosine with custom styling
    Selection.Font.Bold = True
    Selection.TypeText "Example 3: Cosine with Custom Line Width" & vbCrLf
    Selection.Font.Bold = False
    Selection.TypeText "plot(cos(3*x), 0, 2*pi; linewidth=3, title=""Cosine Wave"")" & vbCrLf & vbCrLf
    
    ' Example 4: Exponential function
    Selection.Font.Bold = True
    Selection.TypeText "Example 4: Exponential Decay" & vbCrLf
    Selection.Font.Bold = False
    Selection.TypeText "plot(exp(-x^2), -3, 3; title=""Gaussian"", grid=true, dpi=200)" & vbCrLf & vbCrLf
    
    ' Example 5: Sinc function
    Selection.Font.Bold = True
    Selection.TypeText "Example 5: Sinc Function" & vbCrLf
    Selection.Font.Bold = False
    Selection.TypeText "plot(sin(x)/x, -20, 20; title=""Sinc Function"", backend=""matplotlib"")" & vbCrLf & vbCrLf
    
    ' Example 6: Polynomial
    Selection.Font.Bold = True
    Selection.TypeText "Example 6: Cubic Polynomial" & vbCrLf
    Selection.Font.Bold = False
    Selection.TypeText "plot(x^3 - 3*x^2 + 2*x, -2, 4; title=""Cubic Function"", grid=true)" & vbCrLf & vbCrLf
    
    ' Usage instructions
    Selection.Font.Bold = True
    Selection.TypeText "How to Test:" & vbCrLf
    Selection.Font.Bold = False
    Selection.TypeText "1. Select one of the plot(...) commands above" & vbCrLf
    Selection.TypeText "2. Click WordMat → Plot Selection in the ribbon" & vbCrLf
    Selection.TypeText "3. The plot should be generated and inserted as an image" & vbCrLf & vbCrLf
    
    Selection.Font.Bold = True
    Selection.TypeText "Prerequisites:" & vbCrLf
    Selection.Font.Bold = False
    Selection.TypeText "• For Matplotlib backend: Python with matplotlib and numpy installed" & vbCrLf
    Selection.TypeText "• For Gnuplot backend: Gnuplot executable in system PATH" & vbCrLf & vbCrLf
    
    ' Move cursor to top
    Selection.HomeKey unit:=wdStory
    
    MsgBox "Demo document created successfully!" & vbCrLf & vbCrLf & _
           "You can now select any plot(...) command and test the plotting functionality." & vbCrLf & _
           "Note: You'll need either Python+matplotlib or Gnuplot installed for the backends to work.", _
           vbInformation, "Plot Demo Ready"
    
    Exit Sub
    
DemoError:
    MsgBox "Error creating demo document: " & Err.Description, vbCritical, "Demo Error"
End Sub

' Quick test function that creates a simple plot command and executes it
Public Sub QuickPlotTest()
    On Error GoTo QuickTestError
    
    ' Insert a simple test command
    Selection.TypeText "plot(sin(x), -pi, pi; title=""Quick Test"")"
    Selection.HomeKey unit:=wdLine, Extend:=wdExtend
    
    MsgBox "About to test plot execution with: plot(sin(x), -pi, pi; title=""Quick Test"")" & vbCrLf & vbCrLf & _
           "This will attempt to generate a plot using your default backend." & vbCrLf & _
           "Make sure Python+matplotlib or Gnuplot is installed.", vbInformation, "Quick Plot Test"
    
    ' Execute the plot command
    PlotSelection
    
    Exit Sub
    
QuickTestError:
    MsgBox "Error in QuickPlotTest: " & Err.Description, vbCritical, "Quick Test Error"
End Sub

' Test function to validate backends are available
Public Sub TestBackendsAvailable()
    On Error GoTo BackendTestError
    
    Dim pythonAvailable As Boolean
    Dim gnuplotAvailable As Boolean
    Dim message As String
    
    ' Test Python availability
    pythonAvailable = TestCommandAvailable("python --version")
    
    ' Test Gnuplot availability  
    gnuplotAvailable = TestCommandAvailable("gnuplot --version")
    
    message = "Backend Availability Test Results:" & vbCrLf & vbCrLf
    
    If pythonAvailable Then
        message = message & "✓ Python: Available" & vbCrLf
    Else
        message = message & "✗ Python: Not available" & vbCrLf
    End If
    
    If gnuplotAvailable Then
        message = message & "✓ Gnuplot: Available" & vbCrLf
    Else
        message = message & "✗ Gnuplot: Not available" & vbCrLf
    End If
    
    message = message & vbCrLf
    
    If pythonAvailable Or gnuplotAvailable Then
        message = message & "At least one backend is available. Plot functionality should work!"
    Else
        message = message & "No backends are available. Please install Python+matplotlib or Gnuplot."
    End If
    
    MsgBox message, vbInformation, "Backend Test Results"
    
    Exit Sub
    
BackendTestError:
    MsgBox "Error testing backends: " & Err.Description, vbCritical, "Backend Test Error"
End Sub

' Helper function to test if a command is available
Private Function TestCommandAvailable(command As String) As Boolean
    On Error GoTo CommandTestError
    
    TestCommandAvailable = False
    
    ' This is a simplified test - in a full implementation you would 
    ' use Shell with proper error handling and output capture
    Dim result As Long
    result = Shell(command & " 2>nul", vbHide)
    
    If result > 0 Then
        TestCommandAvailable = True
    End If
    
    Exit Function
    
CommandTestError:
    ' If Shell fails, command is probably not available
    TestCommandAvailable = False
End Function