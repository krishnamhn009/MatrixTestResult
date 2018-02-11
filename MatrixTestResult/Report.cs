using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MatrixTestResult
{
    public class Report
    {
        public void CreateExcelDoc(string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // Adding style
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet.Save();

                // Setting up columns
                Columns columns = new Columns(
                        new Column // Id column
                        {
                            Min = 1,
                            Max = 1,
                            Width = 4,
                            CustomWidth = true
                        },
                        new Column // Name and Birthday columns
                        {
                            Min = 2,
                            Max = 3,
                            Width = 15,
                            CustomWidth = true
                        },
                        new Column // Salary column
                        {
                            Min = 4,
                            Max = 4,
                            Width = 8,
                            CustomWidth = true
                        },
                         new Column // Is Working column
                         {
                             Min = 4,
                             Max = 4,
                             Width = 8,
                             CustomWidth = true
                         });

                worksheetPart.Worksheet.AppendChild(columns);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Employees" };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                List<Employee> employees = Employees.EmployeesList;

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Constructing header
                Row row = new Row();

                row.Append(
                    ConstructCell("Id", CellValues.String, 2),
                    ConstructCell("Name", CellValues.String, 2),
                    ConstructCell("Birth Date", CellValues.String, 2),
                    ConstructCell("Salary", CellValues.String, 2),
                    ConstructCell("IsWorking", CellValues.String, 2));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting each employee
                foreach (var employee in employees)
                {
                    row = new Row();                    
                    row.Append(
                        ConstructCell(employee.Id.ToString(), CellValues.Number, 1),
                        ConstructCell(employee.Name, CellValues.String, 1),
                        ConstructCell(employee.DOB.ToString("yyyy/MM/dd"), CellValues.String, 1),
                        ConstructCell(employee.Salary.ToString(), CellValues.Number, 1),
                        ConstructCell(employee.IsWorking.ToString(), CellValues.String, (uint)(string.Equals(employee.IsWorking, "Yes", StringComparison.OrdinalIgnoreCase) ? 4 : 3)));

                   

                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
            }
        }

        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ));
                        Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })// Index 2 - header
                     { PatternType = PatternValues.Solid }),
                    new Fill(new PatternFill(new ForegroundColor { Rgb = "ff4d4d" }, new BackgroundColor() { Indexed = (UInt32Value)64U })
                      { PatternType = PatternValues.Solid }), ////index 3--red color 
                       new Fill(new PatternFill(new ForegroundColor { Rgb = "00b386" }, new BackgroundColor() { Indexed = (UInt32Value)64U })
                       { PatternType = PatternValues.Solid })
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true }, // header
                    new CellFormat { FontId = 0, FillId = 3, BorderId = 1, ApplyFill = true },
                    new CellFormat { FontId = 0, FillId = 4, BorderId = 1, ApplyFill = true }
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }
    }
}
