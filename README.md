
<h2 align="center">VBA Stock Market Analysis</h2>

## Table of Contents
1. [Overview](#overview)
2. [Code](#code)
3. [Results](#results)
4. [Getting Started](#getting-started)


<p align="center">
  <img src="https://github.com/Xthe23/VBA-challenge/blob/main/ProjectFolder/Images/stockmarket.gif" alt="Stock Market GIF">
</p>

## Overview

This VBA script performs a stock market analysis on multiple worksheets within a workbook. For each worksheet (representing a year), it calculates the yearly change, percent change, and total stock volume for each ticker. It also highlights positive changes in green and negative changes in red. In addition, the script identifies the tickers with the greatest percent increase, greatest percent decrease, and greatest total volume.

## Code

Below is a snippet of the `StockMarketAnalysis` subroutine used in this project. For the complete code, please [click here](https://github.com/Xthe23/VBA-challenge/blob/main/ProjectFolder/StockMarketAnalysis.bas).

```vba
Sub StockMarketAnalysis()
    ' Declare variables for worksheet, row numbers, and ticker information
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    ' ... (rest of the code snippet)
    ' For the full code, visit the link above
End Sub
```
## Results
Here are the results images for the years 2018 and 2019 (in that order):

![Result1](https://github.com/Xthe23/VBA-challenge/blob/main/ProjectFolder/Images/results-2018.png)

![Result2](https://github.com/Xthe23/VBA-challenge/blob/main/ProjectFolder/Images/results-2019.png)

## Getting Started
- Open your Excel workbook.
- Press ALT + F11 to open the VBA editor.
- Insert a new module by clicking Insert > Module.
- Copy and paste the StockMarketAnalysis subroutine into the module.
- Run the subroutine.

  
