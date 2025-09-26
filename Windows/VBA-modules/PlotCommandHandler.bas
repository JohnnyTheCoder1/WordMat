Attribute VB_Name = "PlotCommandHandler"
Option Explicit

' WordMat Plot Command Handler
' Implements command-driven plotting with plot(...) syntax
' Author: WordMat Team

Public Type PlotCommand
    Expression As String
    XMin As String
    XMax As String
    Title As String
    Grid As Boolean
    Width As Integer
    Height As Integer
    Dpi As Integer
    Backend As String
    LineWidth As Double
End Type

Public Type PlotOptions
    DefaultBackend As String
    DefaultWidth As Integer
    DefaultHeight As Integer
    DefaultDpi As Integer
    DefaultGrid As Boolean
    DefaultLineWidth As Double
    PythonPath As String
    GnuplotPath As String
    TimeoutSeconds As Integer
    KeepTempFiles As Boolean
End Type

Private plotOpts As PlotOptions

' Initialize default options
Private Sub InitializePlotOptions()
    With plotOpts
        .DefaultBackend = "matplotlib"
        .DefaultWidth = 800
        .DefaultHeight = 600
        .DefaultDpi = 150
        .DefaultGrid = True
        .DefaultLineWidth = 1.5
        .PythonPath = "python"
        .GnuplotPath = "gnuplot"
        .TimeoutSeconds = 30
        .KeepTempFiles = False
    End With
End Sub

' Main entry point for Plot Selection command
Public Sub PlotSelection()
    On Error GoTo ErrorHandler
    
    ' Initialize options if needed
    If plotOpts.DefaultBackend = "" Then InitializePlotOptions
    
    Dim sel As Selection
    Set sel = Application.Selection
    
    Dim text As String
    text = Trim(sel.Range.text)
    
    If text = "" Then
        ShowPlotError "Please select a plot(...) command to execute." & vbCrLf & _
                     "Example: plot(sin(x), -2*pi, 2*pi)"
        Exit Sub
    End If
    
    If Not LooksLikePlotCommand(text) Then
        ShowPlotError "Selected text does not appear to be a plot(...) command." & vbCrLf & _
                     "Example: plot(sin(x), -2*pi, 2*pi)"
        Exit Sub
    End If
    
    Dim cmd As PlotCommand
    If Not ParsePlotCommand(text, cmd) Then
        Exit Sub ' Error already shown by ParsePlotCommand
    End If
    
    ' Execute the plot command
    ExecutePlotCommand cmd, sel
    
    Exit Sub
    
ErrorHandler:
    ShowPlotError "Error executing plot command: " & Err.Description
End Sub

' Check if text looks like a plot command
Public Function LooksLikePlotCommand(text As String) As Boolean
    text = Trim(text)
    LooksLikePlotCommand = Left(text, 5) = "plot(" And Right(text, 1) = ")"
End Function

' Parse a plot command string into a PlotCommand structure
Public Function ParsePlotCommand(text As String, ByRef cmd As PlotCommand) As Boolean
    On Error GoTo ParseError
    
    ParsePlotCommand = False
    
    ' Remove "plot(" and final ")"
    text = Trim(text)
    If Left(text, 5) <> "plot(" Or Right(text, 1) <> ")" Then
        ShowPlotError "Invalid plot command format"
        Exit Function
    End If
    
    text = Mid(text, 6, Len(text) - 6) ' Remove plot( and )
    
    ' Split by semicolon to separate main params from options
    Dim parts As Variant
    parts = Split(text, ";", 2)
    
    Dim mainPart As String
    mainPart = Trim(parts(0))
    
    ' Parse main parameters (expr, xmin, xmax)
    Dim mainParams As Variant
    mainParams = Split(mainPart, ",")
    
    If UBound(mainParams) < 2 Then
        ShowPlotError "Plot command must have at least 3 parameters: expression, xmin, xmax"
        Exit Function
    End If
    
    ' Set defaults from global options
    With cmd
        .Expression = Trim(mainParams(0))
        .XMin = Trim(mainParams(1))
        .XMax = Trim(mainParams(2))
        .Title = ""
        .Grid = plotOpts.DefaultGrid
        .Width = plotOpts.DefaultWidth
        .Height = plotOpts.DefaultHeight
        .Dpi = plotOpts.DefaultDpi
        .Backend = plotOpts.DefaultBackend
        .LineWidth = plotOpts.DefaultLineWidth
    End With
    
    ' Validate expression for security
    If Not IsValidExpression(cmd.Expression) Then
        ShowPlotError "Invalid or unsafe expression: " & cmd.Expression
        Exit Function
    End If
    
    ' Parse options if present
    If UBound(parts) >= 1 Then
        If Not ParseOptions(Trim(parts(1)), cmd) Then
            Exit Function
        End If
    End If
    
    ParsePlotCommand = True
    Exit Function
    
ParseError:
    ShowPlotError "Error parsing plot command: " & Err.Description
End Function

' Parse option string like "title=""test"", grid=true"
Private Function ParseOptions(optionStr As String, ByRef cmd As PlotCommand) As Boolean
    On Error GoTo OptionsError
    
    ParseOptions = False
    
    If optionStr = "" Then
        ParseOptions = True
        Exit Function
    End If
    
    Dim options As Variant
    options = Split(optionStr, ",")
    
    Dim i As Integer
    For i = 0 To UBound(options)
        Dim option As String
        option = Trim(options(i))
        
        If option <> "" Then
            Dim keyValue As Variant
            keyValue = Split(option, "=", 2)
            
            If UBound(keyValue) < 1 Then
                ShowPlotError "Invalid option format: " & option
                Exit Function
            End If
            
            Dim key As String, value As String
            key = Trim(keyValue(0))
            value = Trim(keyValue(1))
            
            ' Remove quotes from string values
            If Left(value, 1) = """" And Right(value, 1) = """" Then
                value = Mid(value, 2, Len(value) - 2)
            End If
            
            ' Set option values
            Select Case LCase(key)
                Case "title"
                    cmd.Title = value
                Case "grid"
                    cmd.Grid = (LCase(value) = "true")
                Case "width"
                    cmd.Width = CInt(value)
                Case "height"
                    cmd.Height = CInt(value)
                Case "dpi"
                    cmd.Dpi = CInt(value)
                Case "backend"
                    If LCase(value) = "matplotlib" Or LCase(value) = "gnuplot" Then
                        cmd.Backend = LCase(value)
                    Else
                        ShowPlotError "Invalid backend: " & value & ". Use 'matplotlib' or 'gnuplot'"
                        Exit Function
                    End If
                Case "linewidth"
                    cmd.LineWidth = CDbl(value)
                Case Else
                    ShowPlotError "Unknown option: " & key
                    Exit Function
            End Select
        End If
    Next i
    
    ParseOptions = True
    Exit Function
    
OptionsError:
    ShowPlotError "Error parsing options: " & Err.Description
End Function

' Validate expression for security (whitelist approach)
Private Function IsValidExpression(expr As String) As Boolean
    On Error GoTo ValidationError
    
    IsValidExpression = False
    
    ' Allowed functions
    Dim allowedFuncs As String
    allowedFuncs = "sin,cos,tan,asin,acos,atan,exp,log,ln,sqrt,abs"
    
    ' Allowed constants
    Dim allowedConsts As String
    allowedConsts = "pi,π,e,x"
    
    ' Allowed operators and chars
    Dim allowedChars As String
    allowedChars = "+-*/^()0123456789. "
    
    ' Simple validation - check each character/token
    ' This is a basic implementation - a full parser would be better
    Dim i As Integer
    For i = 1 To Len(expr)
        Dim char As String
        char = Mid(expr, i, 1)
        
        If InStr(allowedChars, char) = 0 Then
            ' Check if it's part of a function or constant
            Dim found As Boolean
            found = False
            
            ' Check functions
            Dim funcs As Variant
            funcs = Split(allowedFuncs, ",")
            Dim j As Integer
            For j = 0 To UBound(funcs)
                If i <= Len(expr) - Len(funcs(j)) + 1 Then
                    If Mid(expr, i, Len(funcs(j))) = funcs(j) Then
                        found = True
                        Exit For
                    End If
                End If
            Next j
            
            ' Check constants
            If Not found Then
                Dim consts As Variant
                consts = Split(allowedConsts, ",")
                For j = 0 To UBound(consts)
                    If i <= Len(expr) - Len(consts(j)) + 1 Then
                        If Mid(expr, i, Len(consts(j))) = consts(j) Then
                            found = True
                            Exit For
                        End If
                    End If
                Next j
            End If
            
            If Not found Then
                Exit Function
            End If
        End If
    Next i
    
    IsValidExpression = True
    Exit Function
    
ValidationError:
    IsValidExpression = False
End Function

' Execute the parsed plot command
Private Sub ExecutePlotCommand(cmd As PlotCommand, sel As Selection)
    On Error GoTo ExecuteError
    
    Dim tempPngPath As String
    tempPngPath = Environ("TEMP") & "\wordmat_plot_" & Format(Now, "yyyymmdd_hhnnss") & "_" & Int(Rnd() * 1000) & ".png"
    
    Dim success As Boolean
    success = False
    
    ' Choose backend and render
    If cmd.Backend = "matplotlib" Then
        success = RenderWithMatplotlib(cmd, tempPngPath)
    ElseIf cmd.Backend = "gnuplot" Then
        success = RenderWithGnuplot(cmd, tempPngPath)
    Else
        ShowPlotError "Unknown backend: " & cmd.Backend
        Exit Sub
    End If
    
    If Not success Then
        Exit Sub ' Error already shown by render function
    End If
    
    ' Insert the PNG into Word
    InsertPlotImage tempPngPath, cmd, sel
    
    ' Clean up temp file unless keeping
    If Not plotOpts.KeepTempFiles Then
        If Dir(tempPngPath) <> "" Then
            Kill tempPngPath
        End If
    End If
    
    Exit Sub
    
ExecuteError:
    ShowPlotError "Error executing plot: " & Err.Description
End Sub

' Show plot error in a user-friendly way
Private Sub ShowPlotError(message As String)
    MsgBox message, vbCritical, "WordMat Plot Error"
End Sub

' Render plot using Matplotlib backend
Private Function RenderWithMatplotlib(cmd As PlotCommand, outputPath As String) As Boolean
    On Error GoTo MatplotlibError
    
    RenderWithMatplotlib = False
    
    ' Create Python script
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\wordmat_plot_script_" & Format(Now, "yyyymmdd_hhnnss") & ".py"
    
    Dim pythonScript As String
    pythonScript = GenerateMatplotlibScript(cmd, outputPath)
    
    ' Write script to file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open scriptPath For Output As fileNum
    Print #fileNum, pythonScript
    Close fileNum
    
    ' Execute Python script
    Dim shellCmd As String
    shellCmd = """" & plotOpts.PythonPath & """ """ & scriptPath & """"
    
    Dim result As Long
    result = ExecuteCommandWithTimeout(shellCmd, plotOpts.TimeoutSeconds)
    
    ' Clean up script file unless keeping temp files
    If Not plotOpts.KeepTempFiles Then
        If Dir(scriptPath) <> "" Then
            Kill scriptPath
        End If
    End If
    
    ' Check if output file was created
    If Dir(outputPath) = "" Then
        ShowPlotError "Matplotlib failed to generate plot. Check Python installation and matplotlib package."
        Exit Function
    End If
    
    RenderWithMatplotlib = True
    Exit Function
    
MatplotlibError:
    ShowPlotError "Error with Matplotlib backend: " & Err.Description
End Function

' Generate Python script for matplotlib
Private Function GenerateMatplotlibScript(cmd As PlotCommand, outputPath As String) As String
    Dim script As String
    
    ' Normalize expression: replace π with pi, ^ with **
    Dim expr As String
    expr = cmd.Expression
    expr = Replace(expr, "π", "pi")
    expr = Replace(expr, "^", "**")
    
    ' Normalize range values
    Dim xMin As String, xMax As String
    xMin = Replace(cmd.XMin, "π", "pi")
    xMin = Replace(xMin, "^", "**")
    xMax = Replace(cmd.XMax, "π", "pi")
    xMax = Replace(xMax, "^", "**")
    
    script = "# Generated by WordMat" & vbCrLf
    script = script & "import sys" & vbCrLf
    script = script & "import os" & vbCrLf
    script = script & "import math" & vbCrLf
    script = script & "import numpy as np" & vbCrLf
    script = script & "import matplotlib" & vbCrLf
    script = script & "matplotlib.use('Agg')  # Use non-interactive backend" & vbCrLf
    script = script & "import matplotlib.pyplot as plt" & vbCrLf
    script = script & "" & vbCrLf
    
    script = script & "# Safe namespace with allowed functions" & vbCrLf
    script = script & "ALLOWED_FUNCS = {" & vbCrLf
    script = script & "    'sin': np.sin, 'cos': np.cos, 'tan': np.tan," & vbCrLf
    script = script & "    'asin': np.arcsin, 'acos': np.arccos, 'atan': np.arctan," & vbCrLf
    script = script & "    'exp': np.exp, 'log': np.log, 'ln': np.log, 'sqrt': np.sqrt," & vbCrLf
    script = script & "    'abs': np.abs" & vbCrLf
    script = script & "}" & vbCrLf
    script = script & "ALLOWED_CONSTS = {'pi': math.pi, 'e': math.e}" & vbCrLf
    script = script & "SAFE_NAMES = {**ALLOWED_FUNCS, **ALLOWED_CONSTS}" & vbCrLf
    script = script & "" & vbCrLf
    
    script = script & "def safe_eval(expr, x):" & vbCrLf
    script = script & "    return eval(expr, {'__builtins__': {}}, {**SAFE_NAMES, 'x': x, 'np': np})" & vbCrLf
    script = script & "" & vbCrLf
    
    script = script & "try:" & vbCrLf
    script = script & "    # Parameters" & vbCrLf
    script = script & "    xmin, xmax = float(" & xMin & "), float(" & xMax & ")" & vbCrLf
    script = script & "    N = 2000" & vbCrLf
    script = script & "    x = np.linspace(xmin, xmax, N)" & vbCrLf
    script = script & "    " & vbCrLf
    script = script & "    # Evaluate expression" & vbCrLf
    script = script & "    y = safe_eval('" & expr & "', x)" & vbCrLf
    script = script & "    " & vbCrLf
    script = script & "    # Create plot" & vbCrLf
    script = script & "    fig_width = " & cmd.Width & " / " & cmd.Dpi & vbCrLf
    script = script & "    fig_height = " & cmd.Height & " / " & cmd.Dpi & vbCrLf
    script = script & "    plt.figure(figsize=(fig_width, fig_height), dpi=" & cmd.Dpi & ")" & vbCrLf
    script = script & "    plt.plot(x, y, linewidth=" & cmd.LineWidth & ")" & vbCrLf
    
    If cmd.Grid Then
        script = script & "    plt.grid(True)" & vbCrLf
    End If
    
    If cmd.Title <> "" Then
        script = script & "    plt.title(r'" & EscapeForPython(cmd.Title) & "')" & vbCrLf
    End If
    
    script = script & "    plt.xlabel('x')" & vbCrLf
    script = script & "    plt.ylabel('y')" & vbCrLf
    script = script & "    plt.tight_layout()" & vbCrLf
    script = script & "    plt.savefig(r'" & outputPath & "', dpi=" & cmd.Dpi & ", bbox_inches='tight')" & vbCrLf
    script = script & "    plt.close()" & vbCrLf
    script = script & "    print('Plot saved successfully')" & vbCrLf
    script = script & "" & vbCrLf
    script = script & "except ImportError as e:" & vbCrLf
    script = script & "    print(f'Missing required package: {e}')" & vbCrLf
    script = script & "    sys.exit(1)" & vbCrLf
    script = script & "except Exception as e:" & vbCrLf
    script = script & "    print(f'Error: {e}')" & vbCrLf
    script = script & "    sys.exit(1)" & vbCrLf
    
    GenerateMatplotlibScript = script
End Function

' Escape string for Python
Private Function EscapeForPython(text As String) As String
    EscapeForPython = Replace(text, "'", "\'")
End Function

' Render plot using Gnuplot backend  
Private Function RenderWithGnuplot(cmd As PlotCommand, outputPath As String) As Boolean
    On Error GoTo GnuplotError
    
    RenderWithGnuplot = False
    
    ' Create Gnuplot script
    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\wordmat_plot_script_" & Format(Now, "yyyymmdd_hhnnss") & ".gp"
    
    Dim gnuplotScript As String
    gnuplotScript = GenerateGnuplotScript(cmd, outputPath)
    
    ' Write script to file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open scriptPath For Output As fileNum
    Print #fileNum, gnuplotScript
    Close fileNum
    
    ' Execute Gnuplot script
    Dim shellCmd As String
    shellCmd = """" & plotOpts.GnuplotPath & """ """ & scriptPath & """"
    
    Dim result As Long
    result = ExecuteCommandWithTimeout(shellCmd, plotOpts.TimeoutSeconds)
    
    ' Clean up script file unless keeping temp files
    If Not plotOpts.KeepTempFiles Then
        If Dir(scriptPath) <> "" Then
            Kill scriptPath
        End If
    End If
    
    ' Check if output file was created
    If Dir(outputPath) = "" Then
        ShowPlotError "Gnuplot failed to generate plot. Check gnuplot installation."
        Exit Function
    End If
    
    RenderWithGnuplot = True
    Exit Function
    
GnuplotError:
    ShowPlotError "Error with Gnuplot backend: " & Err.Description
End Function

' Generate Gnuplot script
Private Function GenerateGnuplotScript(cmd As PlotCommand, outputPath As String) As String
    Dim script As String
    
    ' Normalize expression for gnuplot
    Dim expr As String
    expr = cmd.Expression
    expr = Replace(expr, "π", "pi")
    ' Gnuplot supports both ^ and ** so we keep ^
    
    ' Normalize range values
    Dim xMin As String, xMax As String
    xMin = Replace(cmd.XMin, "π", "pi")
    xMax = Replace(cmd.XMax, "π", "pi")
    
    script = "# Generated by WordMat" & vbCrLf
    script = script & "set terminal pngcairo size " & cmd.Width & "," & cmd.Height & " enhanced font 'Arial,12'" & vbCrLf
    script = script & "set output '" & Replace(outputPath, "\", "/") & "'" & vbCrLf
    script = script & "set samples 2000" & vbCrLf
    
    If cmd.Grid Then
        script = script & "set grid" & vbCrLf
    End If
    
    If cmd.Title <> "" Then
        script = script & "set title '" & EscapeForGnuplot(cmd.Title) & "'" & vbCrLf
    End If
    
    script = script & "set xlabel 'x'" & vbCrLf
    script = script & "set ylabel 'y'" & vbCrLf
    script = script & "set xrange [" & xMin & ":" & xMax & "]" & vbCrLf
    script = script & "plot " & expr & " with lines linewidth " & cmd.LineWidth & " title ''" & vbCrLf
    
    GenerateGnuplotScript = script
End Function

' Insert plot image into Word document
Private Sub InsertPlotImage(imagePath As String, cmd As PlotCommand, sel As Selection)
    On Error GoTo InsertError
    
    Dim range As range
    Set range = sel.range
    
    ' Insert the image as an inline shape
    Dim shape As InlineShape
    Set shape = range.InlineShapes.AddPicture( _
        fileName:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        range:=range)
    
    ' Set alternative text for accessibility
    If cmd.Title <> "" Then
        shape.AlternativeText = cmd.Title
    Else
        shape.AlternativeText = "Function plot of " & cmd.Expression & " from " & cmd.XMin & " to " & cmd.XMax
    End If
    
    ' Optionally adjust size (convert pixels to points at 96 DPI)
    If cmd.Width > 0 Then
        shape.Width = cmd.Width * 0.75 ' Convert pixels to points (96 DPI)
    End If
    If cmd.Height > 0 Then
        shape.Height = cmd.Height * 0.75 ' Convert pixels to points (96 DPI)  
    End If
    
    Exit Sub
    
InsertError:
    ShowPlotError "Error inserting image: " & Err.Description
End Sub

' Execute command with timeout (basic implementation using Shell)
Private Function ExecuteCommandWithTimeout(command As String, timeoutSeconds As Integer) As Long
    On Error GoTo CommandError
    
    ' Use Shell to execute command
    ' This is a basic implementation - a more robust version would handle timeout properly
    ExecuteCommandWithTimeout = Shell(command, vbHide)
    
    ' Simple wait - in a full implementation, this would properly handle timeout
    Application.Wait DateAdd("s", 3, Now) ' Wait 3 seconds for command to complete
    
    Exit Function
    
CommandError:
    ExecuteCommandWithTimeout = -1
End Function