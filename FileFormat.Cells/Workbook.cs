using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;

namespace FileFormat.Cells
{
    public class Workbook : IDisposable
    {

        protected internal DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheetDocument;

        protected internal DocumentFormat.OpenXml.Packaging.WorkbookPart workbookpart;

        protected internal DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart;

        protected internal WorkbookStylesPart stylesPart;

        private MemoryStream ms;
        private bool disposedValue;

        private readonly string originalFilePath;

        private BuiltInDocumentProperties _builtinDocumentProperties;

        private uint defaultStyleId;

        public List<Worksheet> Worksheets { get; private set; }

        private StyleUtility styleUtility; // private member

        public Workbook()
        {
            this.ms = new MemoryStream();
            this.spreadsheetDocument = SpreadsheetDocument.Create(this.ms, SpreadsheetDocumentType.Workbook);

            this.workbookpart = this.spreadsheetDocument.AddWorkbookPart();
            this.workbookpart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

            this.Worksheets = new List<Worksheet>(); // Initializing Worksheets list

            // Adding a Worksheet to the Workbook
            var worksheetPart = this.workbookpart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

            // Adding your Worksheet object to Worksheets list
            var newWorksheet = new Worksheet(worksheetPart, worksheetPart.Worksheet);
            this.Worksheets.Add(newWorksheet);

            this.stylesPart = this.spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            this.stylesPart.Stylesheet = new Stylesheet();

            // Append a new worksheet and associate it with the workbook.
            var sheets = this.spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = this.spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            };
            sheets.Append(sheet);

            var workbookStylesPart = workbookpart.WorkbookStylesPart ?? workbookpart.AddNewPart<WorkbookStylesPart>();

            this.styleUtility = new StyleUtility(workbookStylesPart);

            this.defaultStyleId = this.styleUtility.CreateDefaultStyle();

        }

        public Workbook(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentNullException(nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Specified file not found.", filePath);

            this.originalFilePath = filePath;  // store the original file path

            this.ms = new MemoryStream();
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                fs.CopyTo(this.ms);
            }

            this.spreadsheetDocument = SpreadsheetDocument.Open(this.ms, true);
            this.workbookpart = this.spreadsheetDocument.WorkbookPart;

            var workbookStylesPart = workbookpart.WorkbookStylesPart ?? workbookpart.AddNewPart<WorkbookStylesPart>();

            this.styleUtility = new StyleUtility(workbookStylesPart);


            InitializeWorksheets();
        }

        private void InitializeWorksheets()
        {
            this.Worksheets = new List<Worksheet>();

            var sheets = this.workbookpart.Workbook.Sheets.Elements<Sheet>();
            foreach (var sheet in sheets)
            {
                var worksheetPart = (WorksheetPart)(this.workbookpart.GetPartById(sheet.Id));
                var worksheet = worksheetPart.Worksheet;
                var sheetData = worksheet.Elements<SheetData>().FirstOrDefault() ?? new SheetData();
                this.Worksheets.Add(new Worksheet(worksheetPart, worksheet));
            }
        }

        public void UpdateDefaultStyle(string newFontName, double newFontSize, string hexColor)
        {
            // Validate inputs
            if (string.IsNullOrEmpty(newFontName))
                throw new ArgumentNullException(nameof(newFontName));
            if (newFontSize <= 0)
                throw new ArgumentOutOfRangeException(nameof(newFontSize), "Font size must be greater than zero");
            if (string.IsNullOrEmpty(hexColor) || !IsHexColor(hexColor))
                throw new ArgumentException("Invalid hex color", nameof(hexColor));

            // Check if stylesPart exists
            if (workbookpart.WorkbookStylesPart == null)
            {
                stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
            }
            else
            {
                stylesPart = workbookpart.WorkbookStylesPart;
            }

            var stylesheet = stylesPart.Stylesheet;

            // If stylesheet is null, create a new one
            if (stylesheet == null)
            {
                stylesheet = new Stylesheet();
                stylesPart.Stylesheet = stylesheet;
            }

            //var stylesheet = stylesPart.Stylesheet;

            // If stylesheet is null, create a new one
            if (stylesheet == null)
            {
                stylesheet = new Stylesheet();
                stylesPart.Stylesheet = stylesheet;
            }

            // If Fonts collection is null or empty, create a default font
            if (stylesheet.Fonts == null || !stylesheet.Fonts.Any())
            {
                stylesheet.Fonts = new Fonts();
                var defaultFont = new Font();
                stylesheet.Fonts.Append(defaultFont);
            }

            // Assumes the default style is always at index 0.
            //var stylesheet = stylesPart.Stylesheet;
            var font = stylesheet.Fonts.ElementAt(0);
            font.RemoveAllChildren<FontSize>();
            font.RemoveAllChildren<FontName>();

            font.Append(new FontSize() { Val = DoubleValue.FromDouble(newFontSize) });
            font.Append(new FontName() { Val = newFontName });
            font.Append(new Color() { Rgb = new HexBinaryValue() { Value = hexColor } });

            // Save the changes to the stylesheet
            stylesheet.Save();
        }

        private static bool IsHexColor(string color)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(color, "^(#)?([0-9a-fA-F]{3})([0-9a-fA-F]{3})?$");
        }

        public uint DefaultStyleId
        {
            get { return this.defaultStyleId; }
        }

        // Add a public method to create a style, hiding the OpenXML details:
        public uint CreateStyle(string fontName, double fontSize, string hexColor)
        {
            return this.styleUtility.CreateStyle(fontName, fontSize, hexColor);
        }

        

        public Worksheet AddSheet(string sheetName)
        {
            // Create new WorksheetPart and SheetData
            var newWorksheetPart = this.workbookpart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            newWorksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

            // Create a new Worksheet object and add it to Worksheets list
            var newWorksheet = new Worksheet(newWorksheetPart, newWorksheetPart.Worksheet);
            this.Worksheets.Add(newWorksheet);

            // Append a new sheet and associate it with the workbook
            var sheets = this.workbookpart.Workbook.GetFirstChild<Sheets>();
            uint sheetId = (uint)sheets.ChildElements.Count + 1; // Assign the next available SheetId
            var sheet = new Sheet()
            {
                Id = this.workbookpart.GetIdOfPart(newWorksheetPart),
                SheetId = sheetId,
                Name = sheetName
            };
            sheets.Append(sheet);

            return newWorksheet; // Return the newly created Worksheet
        }

        public bool RemoveSheet(string sheetName)
        {
            // Find the sheet in the workbook by its name
            var sheet = workbookpart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

            // If the sheet doesn't exist, return false
            if (sheet == null)
                return false;

            // Retrieve the corresponding worksheet part
            WorksheetPart worksheetPart = (WorksheetPart)(workbookpart.GetPartById(sheet.Id));

            // Remove the sheet from the workbook
            sheet.Remove();

            // Remove the worksheet part from the workbook part
            workbookpart.DeletePart(worksheetPart);

            // Synchronize the Worksheets property with the Sheets of the workbook
            SyncWorksheets();

            // Return true to indicate success
            return true;
        }




        private void SyncWorksheets()
        {
            this.Worksheets = new List<Worksheet>();
            var sheets = this.workbookpart.Workbook.Sheets.Elements<Sheet>();
            foreach (var sh in sheets)
            {
                var wp = (WorksheetPart)(this.workbookpart.GetPartById(sh.Id));
                var ws = wp.Worksheet;
                var sd = ws.Elements<SheetData>().FirstOrDefault() ?? new SheetData();
                this.Worksheets.Add(new Worksheet(wp, ws));
            }
        }




        public void Save()
        {
            if (string.IsNullOrEmpty(this.originalFilePath))
                throw new InvalidOperationException("Original file path is not available. Use SaveAs method to specify a file path.");

            Save(this.originalFilePath); // use the stored original file path
        }


        public void Save(string filePath)
        {
            this.workbookpart.Workbook.Save();
            this.spreadsheetDocument.Close();

            File.WriteAllBytes(filePath, this.ms.ToArray()); // Write the MemoryStream back to the file
        }

        public void Save(Stream stream)
        {

            var clonedDocument = this.spreadsheetDocument.Clone(stream);
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
            stream.Close();

        }

        public BuiltInDocumentProperties BuiltinDocumentProperties
        {
            get
            {
                if (_builtinDocumentProperties != null)
                    return _builtinDocumentProperties;

                _builtinDocumentProperties = new BuiltInDocumentProperties();

                // Access properties through the OpenXml PackageProperties
                var packageProperties = this.spreadsheetDocument.PackageProperties;
                _builtinDocumentProperties.Author = packageProperties.Creator;
                _builtinDocumentProperties.Title = packageProperties.Title;
                _builtinDocumentProperties.CreatedDate = packageProperties.Created.HasValue ? packageProperties.Created.Value : DateTime.MinValue;
                _builtinDocumentProperties.ModifiedBy = packageProperties.LastModifiedBy;
                _builtinDocumentProperties.ModifiedDate = packageProperties.Modified.HasValue ? packageProperties.Modified.Value : DateTime.MinValue;
                _builtinDocumentProperties.Subject = packageProperties.Subject;

                return _builtinDocumentProperties;
            }
            set
            {
                _builtinDocumentProperties = value;

                // Access properties through the OpenXml PackageProperties
                var packageProperties = this.spreadsheetDocument.PackageProperties;
                packageProperties.Creator = value.Author;
                packageProperties.Title = value.Title;
                packageProperties.Created = value.CreatedDate;
                packageProperties.LastModifiedBy = value.ModifiedBy;
                packageProperties.Modified = value.ModifiedDate;
                packageProperties.Subject = value.Subject;

            }
        }



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

