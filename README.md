# WordMat
[![Downloads](https://img.shields.io/github/downloads/Eduap-com/WordMat/total.svg)](https://github.com/Eduap-com/WordMat/releases)
[![Release Version](https://img.shields.io/github/release/Eduap-com/WordMat)](https://github.com/Eduap-com/WordMat/releases/latest)

**The main homepage of the project can be found here:** [Eduap.com](http://www.eduap.com)

*From there you can download the most recent version for Windows and Mac, read the FAQ, see screenshots etc.*

WordMat creates a new ribbon-menu in Word with math functionality. You can do simple and advanced calculations on any math expression entered using the builtin equation editor, plot graphs and much more.

## New: Command-driven Plotting
WordMat now supports command-driven plotting with simple text commands:
```
plot(sin(x), -2π, 2π)
plot(x^2, -5, 5; title="Parabola", grid=true)
```
Select any plot(...) command in your document and click **Plot Selection** to generate and insert a plot image.

Supported backends: Python/Matplotlib and Gnuplot. See [PLOT_FEATURE_README.md](PLOT_FEATURE_README.md) for full details.

This GitHub-site is for people that wants to contribute to the project by reporting bugs, help fix bugs, translate or add new functionality.
See [CONTRIBUTING.md](https://github.com/Eduap-com/WordMat/blob/master/CONTRIBUTING.md)

To Install WordMat read 'How to Install WordMat', or use the free installer on eduap.com.

If you want to learn how to build and change the code read [How To Build WordMat](https://github.com/Eduap-com/WordMat/blob/master/How%20to%20build%20WordMat.md)

WordMat core is open source, released as GNU General public License.

*Mikael Samsøe Sørensen*

