# Excel Report Image Automation Script (Final Version)

## 1. Purpose

This VBA script automates the process of creating a clean, professional inventory report in Microsoft Excel. It is designed to take raw data downloaded from a Google Sheet—where images are represented by an `=IMAGE()` formula—and transform it into a polished report with properly embedded images.

This script works "in-place" for maximum safety and simplicity. It performs the following actions:
*   Identifies the column containing the image formulas (e.g., Column A).
*   Resizes the rows to an appropriate height for viewing images.
*   For each product, it reads the image URL from the formula.
*   It **clears the formula** from the cell.
*   It **inserts the downloaded image** directly into that same cell.
*   The script **does not insert or delete any columns**, which prevents any risk of accidentally removing important data like the "Name" column.

## 2. Requirements

*   Microsoft Excel with the **Developer Tab** enabled.
*   An Excel file saved as a **Macro-Enabled Workbook (`.xlsm`)** to store the script.
*   Report data copied from a Google Sheet where the product image URLs are contained within an `=IMAGE()` formula.

## 3. Step-by-Step Workflow

Follow these steps every time you need to generate a report.

1.  **Generate & Download:** Run your Apps Script in Google Sheets to create the inventory report. Download the result as a Microsoft Excel (`.xlsx`) file.
2.  **Open Files:** Open both the downloaded Excel file and your master Macro-Enabled Template (`.xlsm`) file.
3.  **Copy Data:** In the downloaded file, select and copy all the report data, starting from the header row.
4.  **Paste into Template:** Go to your `.xlsm` file. Click on a fresh sheet and select cell **A1**. Paste the data. The column with the `#NAME?` errors should now be Column A.
5.  **Run the Macro:**
    *   Go to the **Developer** tab in Excel.
    *   Click the **Macros** button.
    *   Select `CreateFinalReport_V2` from the list.
    *   Click **Run**.
6.  **Submit Report:** The script will execute and present a message upon completion. The report is now perfectly formatted with images in Column A and is ready to be saved and emailed.

## 4. The VBA Code

Copy the complete code below and paste it into a new Module within the VBA Editor (`Alt + F11 > Insert > Module`) of your `.xlsm` template file. This is the final and safest version.

```vba
' ======================================================================================
' SCRIPT:           CreateFinalReport_V2
' AUTHOR:           Wesley Quintero
' DATE:             July 21, 2025
' DESCRIPTION:      (SAFE VERSION) Automates the creation of a professional report.
'                   This script works "in-place". It reads a formula from a cell,
'                   clears that same cell, and inserts the final image.
'                   IT DOES NOT INSERT OR DELETE ANY COLUMNS.
' ======================================================================================
Sub CreateFinalReport_V2()
    ' --- User Settings ---
    ' The column that contains the ugly =IMAGE() formulas/errors.
    Const FormulaColumn As String = "A"
    
    ' The first row number that contains actual product data (below the headers).
    Const StartRow As Long = 4
    ' --- End of Settings ---

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False ' Speeds up the macro.

    ' Format the column where the images will be.
    ws.Columns(FormulaColumn).ColumnWidth = 15 ' Set a good width.
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' Check column B for the last row to be safe.

    ' Loop through each product row.
    Dim i As Long
    For i = StartRow To lastRow
        ws.Rows(i).RowHeight = 60 ' Set a uniform row height.
    
        Dim cellContent As String
        Dim targetCell As Range
        Set targetCell = ws.Cells(i, FormulaColumn)
        
        cellContent = targetCell.Formula ' Read the ugly formula from the source.

        If cellContent <> "" And InStr(1, cellContent, "IMAGE(""", vbTextCompare) > 0 Then
            Dim imageUrl As String
            
            ' Extract the first clean URL from the formula text.
            Dim urlBlock As String
            urlBlock = Mid(cellContent, InStr(1, cellContent, """") + 1)
            urlBlock = Left(urlBlock, InStr(1, urlBlock, """") - 1)
            
            If InStr(1, urlBlock, ";") > 0 Then
                imageUrl = Trim(Left(urlBlock, InStr(1, urlBlock, ";") - 1))
            Else
                imageUrl = Trim(urlBlock)
            End If

            ' Proceed only if we found a valid web URL.
            If Left(imageUrl, 4) = "http" Then
            
                ' ** THE SAFE PART: Clear the content of the cell BEFORE adding the image. **
                targetCell.ClearContents

                ' Insert the picture from the web.
                On Error Resume Next ' If a URL is broken, this prevents the script from crashing.
                Dim p As Picture
                Set p = ws.Pictures.Insert(imageUrl)
                
                ' If the picture was inserted successfully, resize and center it.
                If Err.Number = 0 Then
                    With p
                        .ShapeRange.LockAspectRatio = msoTrue
                        .Height = targetCell.Height - 5 ' Resize to fit cell with a small margin.
                        If .Width > targetCell.Width Then
                            .Width = targetCell.Width - 5
                        End If
                        ' Perfectly center the image.
                        .Top = targetCell.Top + (targetCell.Height - .Height) / 2
                        .Left = targetCell.Left + (targetCell.Width - .Width) / 2
                    End With
                Else
                    targetCell.Value = "Image Failed" ' Mark if the URL was bad.
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    
    ws.Columns("A").Select ' Select the image column to finish.
    
    Application.ScreenUpdating = True
    
    MsgBox "Report created successfully! Images have been placed in Column A."
End Sub
```
