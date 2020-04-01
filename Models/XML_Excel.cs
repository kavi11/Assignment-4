using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OpenXML.Models
{
    class XML_Excel
    {
        public static void CreateSpreadsheetWorkbook(String filepath, List<Student> students)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "My Name"
            };
            sheets.Append(sheet);

            WorksheetPart worksheetPart2 = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart2.Worksheet = new Worksheet(new SheetData());
            Sheet sheet2 = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart2),
                SheetId = 2,
                Name = "List of Student"
            };
            sheets.Append(sheet2);
            
            worksheetPart.Worksheet.Save();
       
            SharedStringTablePart shareStringPart;
            shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();

            int row_index = InsertSharedStringItem("Hello, My Name is Kaviraj Singh", shareStringPart);
            Cell cell1 = InsertCellInWorksheet("A", 2, worksheetPart);
            cell1.CellValue = new CellValue(row_index.ToString());
            cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int row_index2 = InsertSharedStringItem("UniqueID", shareStringPart);
            Cell cell2 = InsertCellInWorksheet("A", 1, worksheetPart2);
            cell2.CellValue = new CellValue(row_index2.ToString());
            cell2.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int row_index3 = InsertSharedStringItem("StudentID", shareStringPart);
            Cell cell3 = InsertCellInWorksheet("B", 1, worksheetPart2);
            cell3.CellValue = new CellValue(row_index3.ToString());
            cell3.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int row_index4 = InsertSharedStringItem("FirstName", shareStringPart);
            Cell cell4 = InsertCellInWorksheet("C", 1, worksheetPart2);
            cell4.CellValue = new CellValue(row_index4.ToString());
            cell4.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int row_index5 = InsertSharedStringItem("LastName", shareStringPart);
            Cell cell5 = InsertCellInWorksheet("D", 1, worksheetPart2);
            cell5.CellValue = new CellValue(row_index5.ToString());
            cell5.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int row_index6 = InsertSharedStringItem("DateofBirth", shareStringPart);
            Cell cell6 = InsertCellInWorksheet("E", 1, worksheetPart2);
            cell6.CellValue = new CellValue(row_index6.ToString());
            cell6.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int row_index7 = InsertSharedStringItem("IsMe", shareStringPart);
            Cell cell7 = InsertCellInWorksheet("F", 1, worksheetPart2);
            cell7.CellValue = new CellValue(row_index7.ToString());
            cell7.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int row_index8 = InsertSharedStringItem("Age", shareStringPart);
            Cell cell8 = InsertCellInWorksheet("G", 1, worksheetPart2);
            cell8.CellValue = new CellValue(row_index8.ToString());
            cell8.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            worksheetPart.Worksheet.Save();
            uint rowcount = 2;

            foreach (var student in students)
            {
                String IsMe2 = "0";
                if (student.IsMe == true)
                {
                    IsMe2 = "1";
                }

                int row_index9 = InsertSharedStringItem(student.UniqueID.ToString(), shareStringPart);
                Cell cell9 = InsertCellInWorksheet("A".ToString(), rowcount, worksheetPart2);
                cell9.CellValue = new CellValue(row_index9.ToString());
                cell9.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                int row_index10 = InsertSharedStringItem(student.StudentId.ToString(), shareStringPart);
                Cell cell10 = InsertCellInWorksheet("B".ToString(), rowcount, worksheetPart2);
                cell10.CellValue = new CellValue(row_index10.ToString());
                cell10.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                int row_index11 = InsertSharedStringItem(student.FirstName, shareStringPart);
                Cell cell11 = InsertCellInWorksheet("C".ToString(), rowcount, worksheetPart2);
                cell11.CellValue = new CellValue(row_index11.ToString());
                cell11.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                int row_index12 = InsertSharedStringItem(student.LastName, shareStringPart);
                Cell cell12 = InsertCellInWorksheet("D".ToString(), rowcount, worksheetPart2);
                cell12.CellValue = new CellValue(row_index12.ToString());
                cell12.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                int row_index13 = InsertSharedStringItem(student.DateOfBirthDT.ToString(CultureInfo.InvariantCulture), shareStringPart);
                Cell cell13 = InsertCellInWorksheet("E".ToString(), rowcount, worksheetPart2);
                cell13.CellValue = new CellValue(row_index13.ToString());
                cell13.DataType = new EnumValue<CellValues>(CellValues.SharedString);
 
                int row_index14 = InsertSharedStringItem(IsMe2.ToString(), shareStringPart);
                Cell cell14 = InsertCellInWorksheet("F".ToString(), rowcount, worksheetPart2);
                cell14.CellValue = new CellValue(row_index14.ToString());
                cell14.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                int row_index15 = InsertSharedStringItem(student.Age.ToString(), shareStringPart);
                Cell cell15 = InsertCellInWorksheet("G".ToString(), rowcount, worksheetPart2);
                cell15.CellValue = new CellValue(row_index15.ToString());
                cell15.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                rowcount += 1;
            }

            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        
       
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

    }
}
