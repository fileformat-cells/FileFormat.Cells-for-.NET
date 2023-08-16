using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FileFormat.Cells
{
    /// <summary>
    /// Represents a sheet definition file that contains the sheet data.
    /// </summary>
	public class Worksheet
    {
        /// <value>
        /// An object of the Parent SpreadsheetDocument class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheetDocument;
        /// <value>
        /// An object of the Parent WorkbookPart class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Packaging.WorkbookPart workbookpart;
        /// <value>
        /// An object of the Parent WorksheetPart class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart;
        /// <value>
        /// An object of the Parent Worksheet class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet;
        /// <value>
        /// An object of the Parent Sheets class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Spreadsheet.Sheets sheets;
        private UInt32 sheetID;
        private string ID;
        SheetData sheetData;
        /// <value>
        /// An object of the Parent WorkbookStylesPart class.
        /// </value>
        protected internal WorkbookStylesPart stylesPart;
        /// <value>
        /// An object of the Parent MergeCells class.
        /// </value>
        protected internal MergeCells mergeCells;
        /// <value>
        /// An object of the Parent Stylesheet class.
        /// </value>
        protected internal Stylesheet stylesheet;
        int i = 1;

        /// <summary>
        /// Instantiates a new instance of the Worksheet class.
        /// </summary>
        /// <param name="workbook">An object of the Workbook class.</param>
        public Worksheet(Workbook workbook)
        {
            this.spreadsheetDocument = workbook.spreadsheetDocument;
            this.worksheetPart = workbook.worksheetPart;
            this.worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();
            this.stylesPart = workbook.stylesPart;
            this.stylesheet = this.stylesPart.Stylesheet;
      
            sheetData = new SheetData();
        }

        /// <summary>
        /// Invoke this function to add a Worksheet into a Workbook.
        /// </summary>
        /// <param name="name">A string value.</param>
        public void Add(string name)
        {
            sheetData = new SheetData();
            this.worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();

            this.worksheetPart = this.spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            this.worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();
            this.sheets = this.spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            ID = this.spreadsheetDocument.WorkbookPart.GetIdOfPart(this.worksheetPart);
            sheetID = Convert.ToUInt32(this.spreadsheetDocument.WorkbookPart.Workbook.Sheets.ToList().Count + 1);
            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = ID,
                SheetId = sheetID,
                Name = name
            };
            this.sheets.Append(sheet);

        }

        /// <summary>
        /// Invoke this function to insert text into a Worksheet.
        /// </summary>
        /// <param name="cellRef">A string value.</param>
        /// <param name="rowIndex">An integer value.</param>
        /// <param name="value">A string value.</param>
        /// <param name="format_cell_id">An integer value.</param>
        
        public void insertValue(string cellRef, UInt32 rowIndex, dynamic value, UInt32 format_cell_id)
        {

            Row row = new Row();
            row.Index = rowIndex;
            Cell cell = new Cell();
            cell.setCellReference(cellRef);

            if (value.GetType().ToString() == "System.Int32" || value.GetType().ToString() == "System.Double")
                cell.setNumberDataType();
            if (value.GetType().ToString() == "System.String")
                cell.setStringDataType();
            cell.CellValue(value);
            cell.CellIndex(format_cell_id);

            row.Append(cell);
            sheetData.Append(row);

        }

        /// <summary>
        /// Invoke this function to save data into the Worksheet 
        /// </summary>
        /// <param name="sheetIndex">An integer value.</param>
        public void saveDataToSheet(int sheetIndex)
        {
            if (this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].Worksheet.ChildElements.ToList().Count == 1)
            {

                var obj = this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].Worksheet.ChildElements.ToList()[0];
                this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].Worksheet.ChildElements.ToList()[0].Remove();
                this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].Worksheet.Append(sheetData.sheetData);
                this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].Worksheet.Append(obj);
            }
            else
            {

                worksheet.Append(sheetData.sheetData);
                this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].Worksheet = worksheet;
            }
            // Add the MergeCells element to the worksheet, if it exists.
            if (mergeCells != null)
            {
                worksheet.Append(mergeCells);
            }

            // Save the worksheet data
            worksheet.Save();
            


        }


        private static List<UInt32> AddStyle(ref Stylesheet stylesheet)
        {
            UInt32 Fontid = 0, Fillid = 0, Cellformatid = 0;

            Font font = new Font(new FontSize() { Val = 36 },
                                 new BackgroundColor() { Rgb = HexBinaryValue.FromString("FFFFFF") });

            if (stylesheet.Fonts == null)
            {
                stylesheet.Fonts = new Fonts();
                stylesheet.Fonts.Count = 1;
            } else
            {
                stylesheet.Fonts.Count++;
            }
            
            stylesheet.Fonts.AppendChild(font);
            
            Fontid = stylesheet.Fonts.Count;

            PatternFill pfill = new PatternFill() { PatternType = PatternValues.Solid };
            pfill.BackgroundColor = new BackgroundColor() { Rgb = HexBinaryValue.FromString("70AD47") };

            if (stylesheet.Fills == null)
            {
                stylesheet.Fills = new Fills();
                stylesheet.Fills.Count = 1;
            }
            else
            {
                stylesheet.Fills.Count++;
            }
            stylesheet.Fills.Append(new Fill() { PatternFill = pfill });
            
            Fillid = stylesheet.Fills.Count;

           

            CellFormat cellFormat = new CellFormat()
            {
                FontId = Fontid,
                FillId = Fillid,
                ApplyFill = true,
                Alignment = new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Center,
                    Vertical = VerticalAlignmentValues.Center
                }
            };
            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new CellFormats();
                stylesheet.CellFormats.Count = 1;
            }
            else
            {
                stylesheet.CellFormats.Count++;
            }
            stylesheet.CellFormats.AppendChild(cellFormat);
            
            Cellformatid = stylesheet.CellFormats.Count++;

            return new List<uint>() { Fontid, Fillid, Cellformatid };

        }

        /// <summary>
        /// Insert a new cell style into the spreadsheet's styles part.
        /// </summary>
        /// <param name="cellStyle">The <see cref="CellStyle"/> contains the desired font family, size, and cell color.</param>
        /// <returns>The index of the inserted style within the stylesheet's cell formats.</returns>


        public UInt32 insertStyle(CellStyle cellStyle)
        {
            
            // blank font list
            if (this.stylesPart.Stylesheet.Fonts == null)
            {
                this.stylesPart.Stylesheet.Fonts = new Fonts();
                this.stylesPart.Stylesheet.Fonts.AppendChild(new Font());
            }            

            // Create a font with the desired family and size
            Font font1 = new Font();
            font1.FontSize = new FontSize() { Val = cellStyle.FontSize };
            font1.FontName = new FontName() { Val = cellStyle.FontName };            
            this.stylesPart.Stylesheet.Fonts.AppendChild(font1);
            

            // create fills
            stylesPart.Stylesheet.Fills = new Fills();

            // create a solid red fill
            var solidRed = new PatternFill() { PatternType = PatternValues.Solid };
            solidRed.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString(cellStyle.CellColor) };
            solidRed.BackgroundColor = new BackgroundColor { Indexed = 64 };

            this.stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            this.stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            this.stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = solidRed });
            this.stylesPart.Stylesheet.Fills.Count = 3;

            // blank border list
            this.stylesPart.Stylesheet.Borders = new Borders();
            this.stylesPart.Stylesheet.Borders.Count = 1;
            this.stylesPart.Stylesheet.Borders.AppendChild(new Border());

            // blank cell format list
            this.stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
            this.stylesPart.Stylesheet.CellStyleFormats.Count = 1;
            this.stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            // cell format list
            if (this.stylesPart.Stylesheet.CellFormats == null)
            {
                this.stylesPart.Stylesheet.CellFormats = new CellFormats();

                // empty one for index 0, seems to be required
                this.stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());

            }
            
            // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
            this.stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = ( (uint) this.stylesPart.Stylesheet.Fonts.ChildElements.Count - 1 ) , BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });

            this.stylesPart.Stylesheet.Save();

            return ((uint)this.stylesPart.Stylesheet.CellFormats.ChildElements.Count - 1);
        }


        // Given a cell name, parses the specified cell to get the column name.
        private static string GetColumnName(string cellName)
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        // Given a cell name, parses the specified cell to get the row index.
        private static uint GetRowIndex(string cellName)
        {
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        
        
        /// <summary>
        /// Merge two cells by range (e.g., "A1" to "B2")
        /// </summary>
        /// <param name="startCellRef">A string value.</param>
        /// <param name="endCellRef">A string value.</param>
        public void MergeCells(string startCellRef, string endCellRef)
        {

            if (worksheet.Elements<MergeCells>().Count() > 0)
                mergeCells = worksheet.Elements<MergeCells>().First();
            else
            {
                mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.  
                // Insert a MergeCells object into the specified position.  
                if (worksheet.Elements<CustomSheetView>().Any())
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                }
                else
                {
                    var sheetData = worksheetPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().FirstOrDefault();
                    if (sheetData != null)
                    {
                        worksheetPart.Worksheet.InsertAfter(mergeCells, sheetData);
                    }
                }

            }

            // Create the merged cell and append it to the MergeCells collection.  
            MergeCell mergeCell = new MergeCell()
            {
                Reference =
                new StringValue(startCellRef + ":" + endCellRef)
            };
            mergeCells.Append(mergeCell);

        }


    }
}

