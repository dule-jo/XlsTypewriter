
## Summary
ExcelTypewriter is .NET library for writing Excel files built on top of [ClosedXML](https://github.com/ClosedXML/ClosedXML). It aims to provide intuitive and user friendly interface, simulating typewriter. User doesn't have to worry about cell names. Instead of writing cell names him/herself (e.g. "C3"), user navigates excel file using functions Print, Skip and NextRow.
## Installation
Package can be installed as Nuget package.
```
dotnet add package XlsTypewriter
```
## What can you do with this?
XlsTypewriter allows you to write Excel file without keeping track of cell names. 

**Example:**
```  
using XlsTypewriter;  
  
using var workbook = new Workbook();  
var worksheet = workbook.AddWorksheet("Sheet1");

for (var i = 1; i < 4; i++) {
	for (var j = 1; j < 3; j++) { 
		worksheet.Print(i * j);
	}
	worksheet.NewRow();
}

workbook.SaveAs("Example.xlsx");
```
## Basic functions
```
worksheet.Print("Hello World"); // print text in current cell
worksheet.NewRow(); // move to next row
worksheet.NewRow(2); // move to next 2 rows
worksheet.NewRowKeepColumn(); // move to next row, but keep current column
worksheet.NewRowKeepColumn(2); // move to next 2 rows, but keep current column
worksheet.Skip(); // move to next cell in current row
worksheet.Skip(2); // move to next 2 cells in current row
worksheet.PrintAndGoDown("Hello World"); // print text in current cell and move to next row
```
## Merging cells
```
worksheet.Merge(rows:1, columns:2); // merge current cell with next cell in row, getting 2*1
worksheet.Merge(2, 2); // merge current cell with cells in next row and in next column, getting 2x2 cell
```
## Style
Style can be set for current cell or for range of cells.
```
var style = worksheet.GetStyle();
style.Font.Bold = true;
worksheet.SetStyle(style, 1, 2); // set style for cell in 2nd column and 1st row
```
Style can be set using predefined styles. Here are some examples:
```
worksheet.SetBoxBorders(XLBorderStyleValues.Thick); // set thick border around current cell
worksheet.SetBordersColor(XLColor.Blue, 1, 2); // set blue color for borders current cell and next cell in row
worksheet.SetVerticalBorders(XLBorderStyleValues.Thin, 2, 1); // set thin vertical borders for current cell and next cell in column
worksheet.SetHorizontalBorders(XLBorderStyleValues.Medium);
worksheet.SetFontColor(XLColor.Amber);
worksheet.SetBackgroundColor(XLColor.Bistre);
worksheet.SetFontName("Arial");
worksheet.SetFontSize(12);
worksheet.SetFontItalic();
worksheet.SetFontUnderline();
worksheet.SetVerticalAlignment(XLAlignmentVerticalValues.Center);
worksheet.SetHorizontalAlignment(XLAlignmentHorizontalValues.Center);
worksheet.SetBordersColor(XLColor.Blue);
worksheet.SetTopBorder(XLBorderStyleValues.Thin);
worksheet.SetBottomBorder(XLBorderStyleValues.Thin);
worksheet.SetLeftBorder(XLBorderStyleValues.Thin);
worksheet.SetRightBorder(XLBorderStyleValues.Thin);
```
Style can be set for whole file
```
var wsstyle = worksheet.GetStyle();
wsstyle.Font.Bold = true;
worksheet.SetStyleToWorksheet(wsstyle);
```
Style can be set for whole row and column
```
worksheet.SetRowStyle(style);  // set style for current row
worksheet.SetRowStyle(style, 3); // set style for current and next 2 rows

worksheet.SetColumnStyle(style); // set style for current column
worksheet.SetColumnStyle(style, 3); // set style for current and next 2 columns
```
## Navigation
Beside regular Print, Skip and NewRow navigating through file, you can use 
```
GoTo(int column, int row) // go to cell with given column and row
GoToStart() // go to begin of file, or "A1" cell
GoToEnd() // go to last used cell in file
 ```
Position can be saved and restored using position stack
```
worksheet.PushPosition(); // save current position
worksheet.PopPosition(); // restore last saved position
worksheet.PurgePositions(); // clear position stack
```
## Adjusting column width and row height
```
worksheet.AdjustColumnWidth(); // adjust width of all used columns
worksheet.AdjustCurrentColumnWidth(); // adjust width of current column

worksheet.AdjustRowHeight(); // adjust height of all used rows
worksheet.AdjustCurrentRowHeight(); // adjust height of current row
```
## Formulas
*TO DO*
## Properties
```
workbook.Properties.Title = "Hello World";
workbook.CustomProperties.Add("Author2", "John Doe");
```
## Hide and unhide
```
worksheet.HideColumn(); // hide current column
worksheet.HideRow();  // hide current row
```
## Page setup
Page setup can be accessed through worksheet.PageSetup. Here are some examples:
```
worksheet.PageSetup.Header.Left.AddText("Created with XlsTypewriter");
worksheet.PageSetup.Margins.Top = 1;
worksheet.PageSetup.VerticalDpi = 600;
```

## Freeze rows and columns
```
worksheet.FreezeRows() // freeze current row
worksheet.FreezeColumns() // freeze current column
```

## Save as stream
```
using var stream = new MemoryStream();
workbook.SaveAs(stream);
```