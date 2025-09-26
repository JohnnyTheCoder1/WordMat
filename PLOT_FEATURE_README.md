# WordMat Plot Command Feature

## Overview
This document describes the new command-driven plotting feature for WordMat that allows users to create plots using text commands like `plot(sin(x), -2π, 2π)`.

## Usage

### Basic Syntax
```
plot(expression, xmin, xmax)
plot(expression, xmin, xmax; option=value, option=value)
```

### Examples
```
plot(sin(x), -2π, 2π)
plot(x^2, -5, 5; title="Parabola", grid=true, width=600, height=400, dpi=150)
plot(cos(3*x), 0, 10; backend="gnuplot", linewidth=2)
plot(sin(x)/x, -20, 20; title="Sinc Function")
```

### How to Use
1. Type a plot command in your Word document
2. Select the plot command text
3. Click **WordMat → Plot Selection** in the ribbon, or use keyboard shortcut **Alt+W, P**
4. The plot will be generated and inserted as an image at the cursor location

### Supported Functions
- **Trigonometric:** sin, cos, tan, asin, acos, atan
- **Exponential/Logarithmic:** exp, log, ln
- **Other:** sqrt, abs
- **Constants:** pi (or π), e
- **Operators:** +, -, *, /, ^, ()

### Options
- **title**: Plot title (string)
- **grid**: Show grid lines (true/false)
- **width**: Image width in pixels (integer)
- **height**: Image height in pixels (integer) 
- **dpi**: Image resolution (integer)
- **backend**: Rendering backend ("matplotlib" or "gnuplot")
- **linewidth**: Line thickness (number)

### Backends
- **Matplotlib** (default): Requires Python with matplotlib and numpy packages
- **Gnuplot**: Requires gnuplot executable

## Implementation Files

### VBA Modules Added/Modified
- **PlotCommandHandler.bas**: Core plotting functionality
- **PlotCommandTests.bas**: Test suite for validation
- **RibbonSubs.bas**: Added `Rib_PlotSelection` callback
- **ModuleSettings.bas**: Added `PlotSelection` to KeybShortcut enum
- **ModuleKeyboardShortcuts.bas**: Added keyboard shortcut handling

### Key Functions
- **PlotSelection()**: Main entry point, parses selected text and executes plot
- **ParsePlotCommand()**: Parses plot(...) syntax with security validation
- **RenderWithMatplotlib()**: Python/Matplotlib backend implementation
- **RenderWithGnuplot()**: Gnuplot backend implementation
- **InsertPlotImage()**: Inserts generated PNG into Word document

## Security
- Expression validation uses strict whitelist of allowed functions and operators
- No arbitrary code execution - only mathematical expressions are allowed
- Temporary files are cleaned up automatically (unless debug mode is enabled)

## Configuration
Default settings can be customized:
- Default backend (matplotlib/gnuplot)
- Default image dimensions (800x600)
- Default DPI (150)
- Backend executable paths
- Timeout settings

## Ribbon Integration
To add the Plot Selection button to the WordMat ribbon:

1. Open WordMat.dotm in Word
2. Go to File → Options → Customize Ribbon
3. Add a new button with:
   - **Action**: Rib_PlotSelection 
   - **Label**: "Plot Selection"
   - **Group**: WordMat → Graphs (or create new Commands group)

Alternatively, the ribbon XML can be modified to include:
```xml
<button id="plotSelection" label="Plot Selection" onAction="Rib_PlotSelection" 
        imageMso="ChartInsert" size="normal"/>
```

## Testing
Run tests with:
```vba
' Basic parsing and validation tests
RunPlotTests

' Interactive test with sample plot
TestPlotExample
```

## Dependencies
- **For Matplotlib backend**: Python with matplotlib, numpy packages
- **For Gnuplot backend**: Gnuplot executable in system PATH
- Both backends are optional - users can choose their preferred one

## Future Enhancements
- 3D plotting support
- Multi-axis plots
- Parametric plots
- Plot style customization
- Interactive plot editing
- Export to other formats (SVG, PDF)