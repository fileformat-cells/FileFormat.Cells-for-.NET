using System;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FileFormat.Cells.Properties;

namespace FileFormat.Cells
{
    /// <summary>
    /// Represents a root object to create an Excel spreadsheet.
    /// </summary>
	public class Workbook : IDisposable
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
        /// An object of the Parent worksheetPart class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart;
        /// <value>
        /// An object of the Parent WorkbookStylesPart class.
        /// </value>
        protected internal WorkbookStylesPart stylesPart;

        private MemoryStream ms;
        private bool disposedValue;

        /// <summary>
        /// Instantiate a new instance of the Workbook class.
        /// </summary>
        public Workbook()
        {
            this.ms = new MemoryStream();
            this.spreadsheetDocument = SpreadsheetDocument.Create(this.ms, SpreadsheetDocumentType.Workbook);

            this.workbookpart = this.spreadsheetDocument.AddWorkbookPart();
            this.workbookpart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

            this.worksheetPart = this.workbookpart.AddNewPart<WorksheetPart>();
            this.worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();

            this.stylesPart = this.spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            this.stylesPart.Stylesheet = new Stylesheet();


            Sheets sheets = this.spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new Sheet()
            { Id = this.spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

        }

        /// <summary>
        /// Applies the specified font style to the workbook.
        /// </summary>
        /// <param name='fontName'>The name of the font to be applied.</param>
        /// <param name="fontSize">The size of the font to be applied.</param>

        public void ApplyFontStyle(string fontName, int fontSize)
        {
            // Get the WorkbookStylesPart
            WorkbookStylesPart stylesPart = workbookpart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            if (stylesPart == null)
            {
                stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
            }

            // Create the font with the desired font name and size
            Font font = new Font();
            font.FontSize = new FontSize() { Val = fontSize };
            font.FontName = new FontName() { Val = fontName };

            // Add the font to the Fonts collection in the Stylesheet
            stylesPart.Stylesheet.Fonts = new Fonts();
            stylesPart.Stylesheet.Fonts.AppendChild(font);
            
        }


        /// <summary>
        /// Create an object of the Workbook class that opens an existing Excel document from a file.
        /// </summary>
        /// <param name="filePath">String value that represents the document name.</param>
        public Workbook(string filePath)
        {
            this.ms = new MemoryStream();
            FileStream fs = new FileStream(filePath, FileMode.Open);
            fs.CopyTo(this.ms);
            
            this.spreadsheetDocument = SpreadsheetDocument.Open(this.ms, true);
            this.workbookpart = this.spreadsheetDocument.WorkbookPart;
           
        }

        /// <summary>
        /// Create an object of the Workbook class that opens an existing Excel document from a stream.
        /// </summary>
        /// <param name="stream">An object of the Stream class.</param>
        public Workbook(Stream stream)
        {
            this.ms = new MemoryStream();
            stream.CopyTo(this.ms);
            this.spreadsheetDocument = SpreadsheetDocument.Open(this.ms, true);
            this.workbookpart = this.spreadsheetDocument.WorkbookPart;
        }

        /// <summary>
        /// Invoke this method to save the document to a file. 
        /// </summary>
        /// <param name="filePath">String value represents the document name.</param>
        public void Save(string filePath)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
            {
                for (int i = 0; i < this.spreadsheetDocument.WorkbookPart.Workbook.Sheets.ToList().Count; i++) {
                    if (this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[i].Worksheet.ChildElements.FirstOrDefault() != null && this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[i].Worksheet.ChildElements.FirstOrDefault().ToString() != "DocumentFormat.OpenXml.Spreadsheet.SheetData")
                    {

                        var obj = this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[i].Worksheet.ChildElements.ToList()[0];
                        this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[i].Worksheet.ChildElements.ToList()[0].Remove();

                        this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[i].Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
                        this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[i].Worksheet.Append(obj);

                    }
                    else if(this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[i].Worksheet.ChildElements.ToList().Count == 0)
                    {
                        this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[i].Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.SheetData());
                    }
                    

                }
                this.workbookpart.Workbook.Save();
                this.spreadsheetDocument.Close();
             
                this.ms.WriteTo(fileStream);
            }
        }

        /// <summary>
        /// Invoke this method to save the document to a stream. 
        /// </summary>
        /// <param name="stream">An object of the Stream class.</param>
        public void Save(Stream stream)
        {
           
            var clonedDocument = this.spreadsheetDocument.Clone(stream);
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
            stream.Close();
        
        }

        /// <summary>
        /// This property returns the number of worksheets existing in an Excel document.
        /// </summary>
        /// <value>
        /// An integer value.
        /// </value>
        public int Worksheets
        {
            get { return this.spreadsheetDocument.WorkbookPart.Workbook.Sheets.ToList().Count; }
        }

        /// <summary>
        /// Invoke this method to delete a worksheet from an Excel document. 
        /// </summary>
        /// <param name="sheetName">String value represents the worksheet name.</param>
        /// <returns>A string value. </returns>
        public string DeleteWorksheet(string sheetName)
        {
            try
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                int sheetCount = workbookPart.Workbook.Sheets.Count();

                // Get the SheetToDelete from workbook.xml  
                var theSheet = workbookPart.Workbook.Descendants<Sheet>()
                                            .FirstOrDefault(s => s.Name == sheetName);

                if (sheetCount <= 1 || theSheet == null)
                {
                    return $"sheet count from the excel file should be greater than 1 (or) the sheet doesn't exist..";
                }

                // Remove the sheet reference from the workbook.  
                var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(theSheet.Id));
                theSheet.Remove();

                // Delete the worksheet part.  
                workbookPart.DeletePart(worksheetPart);

                // Save the workbook.  
                workbookPart.Workbook.Save();
                spreadsheetDocument.Save();

                return $"{sheetName} deleted successfully..";
            }
            catch (Exception ex)
            {
                return $"ERROR: {ex.Message}";
            }
        }

        /// <summary>
        /// Invoke this method to read the value of a particular cell. 
        /// </summary>
        /// <param name="sheetName">String value represents the worksheet name.</param>
        /// <param name="cellRef">String value represents the address of the cell.</param>
        /// <returns>A sttring value. </returns>
        public string GetCellValue(string sheetName, string cellRef)
        {

            Sheet currentSheet = this.spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().
             Where(s => s.Name == sheetName).FirstOrDefault();

            if (currentSheet == null)
                return "No Worksheet found";

            WorksheetPart wsPart = (WorksheetPart)(this.spreadsheetDocument.WorkbookPart.GetPartById(currentSheet.Id));
            DocumentFormat.OpenXml.Spreadsheet.Cell currentCell = wsPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().
              Where(c => c.CellReference == cellRef).FirstOrDefault();

            if (currentCell == null)
                return null;
            return currentCell.InnerText;
        }

        /// <summary>
        /// This method removes a cell's value in a Worksheet.
        /// </summary>
        /// <param name="sheetName">String value represents the worksheet name.</param>
        /// <param name="colName">The string value represents the name of the column.</param>
        /// <param name="rowIndex">An integer value represents the row.</param>
        public void DeleteTextFromCell(string sheetName, string colName, uint rowIndex)
        {

            IEnumerable<Sheet> sheets = this.spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }
            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);

            // Get the cell at the specified column and row.
            DocumentFormat.OpenXml.Spreadsheet.Cell cell = GetSpreadsheetCell(worksheetPart.Worksheet, colName, rowIndex);
            if (cell == null)
            {
                // The specified cell does not exist.
                return;
            }

            cell.Remove();

        }

        private DocumentFormat.OpenXml.Spreadsheet.Cell GetSpreadsheetCell(DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet, string columnName, uint rowIndex)
        {
            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Row> rows = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Where(r => r.RowIndex == rowIndex);
            if (rows.Count() == 0)
            {
                // A cell does not exist at the specified row.
                return null;
            }

            IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Cell> cells = rows.First().Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
            if (cells.Count() == 0)
            {
                // A cell does not exist at the specified column, in the specified row.
                return null;
            }

            return cells.First();
        }

        /// <summary>
        /// It returns custom built-in document properties.
        /// </summary>
        /// <returns>An object of document properties. </returns>
        public BuiltInDocumentProperties BuiltinDocumentProperties
        {
            get
            {
                BuiltInDocumentProperties prop = new BuiltInDocumentProperties();

                using (var package = Package.Open(this.ms))
                {
                    prop.Author = package.PackageProperties.Creator;
                    prop.Title = package.PackageProperties.Title;
                    if (package.PackageProperties.Created != null)
                        prop.CreatedDate = (DateTime)package.PackageProperties.Created;
                    prop.ModifiedBy = package.PackageProperties.LastModifiedBy;
                    if (package.PackageProperties.Modified != null)
                        prop.ModifiedDate = (DateTime)package.PackageProperties.Modified;
                }
                return prop;
            }
        }

        /// <summary>
        /// This method releases unmanaged resources. 
        /// </summary>
        /// <param name="disposing">A boolean value.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    this.spreadsheetDocument.Dispose();
                    this.ms.Dispose();
                }


                disposedValue = true;
            }
        }

        /// <summary>
        /// This method releases unmanaged resources. 
        /// </summary>
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}

