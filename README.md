# Excel VBA
======

## Overview
The Excel VBA macros are designed to simplify and automate common tasks in Excel. 
This project demonstrates the power of VBA for streamlining workflows, making it an excellent showcase for learning or portfolio purposes.

## Features
This repository includes a variety of VBA scripts to perform tasks such as:

### 1. Data Formatting:
    * Auto-formatting columns (e.g., text alignment, font changes, borders).
    * Converting data to tables with consistent styling.

### 2. Data Cleaning:
   * Removing duplicates.
   * Handling blank rows and columns.
   * Trimming unnecessary spaces from cell values.

### 3. Automation:
   * Automatically generating reports based on a template.
   * Batch renaming worksheets.
   * Sorting and filtering data dynamically.

### 4. Navigation and Interaction:
   * Adding buttons and user forms for easier interaction.
   * Navigating between worksheets with shortcut macros.

## Getting Started
### Prerequisites
   * Microsoft Excel (2016 or later recommended).
   * Basic understanding of VBA (Visual Basic for Applications).

### Installation
1. Enable the Developer Tab:
   * Open Excel.
   * Go to `File` > `Options` > `Customize Ribbon`.
   * Check the Developer option and click `OK`.

2. Access VBA Editor:
   * Press `Alt + F11` to open the VBA editor.

3. Import the Code:
   * Download this repository as a `.zip` or clone it.
   * Open the Excel workbook where you'd like to add the macros.
   * Import the `.bas` files into your project via the VBA editor.

## How to Use
   * Open the workbook containing the macros.
   * Run the desired macro via the VBA editor (`Alt + F8`) or assign it to a button in Excel.
   * Follow the instructions within each macro for customization.


## Code Examples
Here are a few examples of the VBA scripts included:


### Auto-Format Columns

```vbnet
Sub AutoFormatColumns()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    With ws.UsedRange
        .Font.Name = "Calibri"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
End Sub
```

### Remove Duplicates

```vbnet
Sub RemoveDuplicates()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ws.UsedRange.RemoveDuplicates Columns:=Array(1), Header:=xlYes
End Sub
```

## Why Use VBA?
VBA (Visual Basic for Applications) is a powerful tool built into Microsoft Office products that allows users to automate repetitive tasks, manipulate data, and build interactive tools directly within Excel.


## Contributing
Contributions are welcome! Feel free to fork the repository, make your changes, and submit a pull request. Suggestions for additional macros are also appreciated.

## Contact
For questions, suggestions, or issues, feel free to reach out via GitHub or email at `cliffsiu39@gmail.com` .
