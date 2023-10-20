using System;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FileFormat.Cells
{
    /// <summary>
    /// Represents a worksheet within an Excel file, providing methods to manipulate its content.
    /// </summary>
    public sealed class Worksheet
    {
        private WorksheetPart _worksheetPart;
        private SheetData _sheetData;

        /// <summary>
        /// Gets the indexer for cells within the worksheet.
        /// </summary>
        public CellIndexer Cells { get; }

        /// <summary>
        /// Initializes a new instance of the Worksheet class.
        /// </summary>
        /// <param name="worksheetPart">The worksheet part of the document.</param>
        /// <param name="worksheet">The underlying OpenXML worksheet instance.</param>
        private Worksheet(WorksheetPart worksheetPart, DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet)
        {
            _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));

            _sheetData = worksheet?.Elements<SheetData>().FirstOrDefault()
                         ?? throw new InvalidOperationException("SheetData not found in the worksheet.");

            // Initialize the Cells property
            this.Cells = new CellIndexer(this);
        }

        public class WorksheetFactory
        {
            public static Worksheet CreateInstance(WorksheetPart worksheetPart, DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet)
            {
                return new Worksheet(worksheetPart, worksheet);
            }
        }

        /// <summary>
        /// Gets or sets the name of the worksheet.
        /// </summary>
        public string Name
        {
            get
            {
                if (_worksheetPart == null)
                    throw new InvalidOperationException("WorksheetPart is null.");

                var workbookPart = _worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
                if (workbookPart == null)
                    throw new InvalidOperationException("WorkbookPart not found as a parent.");

                var id = workbookPart.GetIdOfPart(_worksheetPart);
                if (string.IsNullOrEmpty(id))
                    throw new InvalidOperationException("ID is null or empty.");

                var sheet = workbookPart.Workbook.Sheets.Cast<Sheet>().FirstOrDefault(s => s.Id.Value.Equals(id));
                if (sheet == null)
                    throw new InvalidOperationException("Sheet not found with the specified ID.");

                return sheet.Name;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    throw new ArgumentException("Sheet name cannot be null or empty", nameof(value));

                var workbookPart = _worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
                var sheet = workbookPart?.Workbook.Sheets.Cast<Sheet>()
                              .FirstOrDefault(s => s.Id.Value.Equals(workbookPart.GetIdOfPart(_worksheetPart)));

                if (sheet != null)
                    sheet.Name = value;
            }
        }


        /// <summary>
        /// Retrieves a cell based on its reference.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation.</param>
        /// <returns>The cell at the specified reference.</returns>
        public Cell GetCell(string cellReference)
        {
            // This logic used to be in your indexer
            return new Cell(GetOrCreateCell(cellReference), _sheetData);
        }

        /// <summary>
        /// Adds an image to the worksheet.
        /// </summary>
        /// <param name="image">The image to be added.</param>
        /// <param name="startRowIndex">The starting row index.</param>
        /// <param name="startColumnIndex">The starting column index.</param>
        /// <param name="endRowIndex">The ending row index.</param>
        /// <param name="endColumnIndex">The ending column index.</param>
        public void AddImage(Image image, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            if (image == null) throw new ArgumentNullException(nameof(image));

            // Assuming you have a working constructor or factory method for ImageHandler
            var imgHandler = new ImageHandler(_worksheetPart);
            imgHandler.Add(image.Path, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        }

        /// <summary>
        /// Extracts images from the worksheet.
        /// </summary>
        /// <returns>A list of images present in the worksheet.</returns>
        public List<Image> ExtractImages()
        {
            List<Image> imagePartsCollection = new List<Image>();

            if (this._worksheetPart.DrawingsPart == null)
                return imagePartsCollection; // Return an empty list instead of null

            foreach (var part in this._worksheetPart.DrawingsPart.ImageParts)
            {
                var stream = part.GetStream();
                var extension = GetImageExtension(part.ContentType);
                imagePartsCollection.Add(new Image(stream, extension));
            }
            return imagePartsCollection;
        }

        /// <summary>
        /// Sets the height of a specific row.
        /// </summary>
        /// <param name="rowIndex">The index of the row.</param>
        /// <param name="height">The desired height.</param>
        public void SetRowHeight(uint rowIndex, double height)
        {
            var row = GetOrCreateRow(rowIndex);
            row.Height = height;
            row.CustomHeight = true;
        }

        /// <summary>
        /// Sets the width of a specific column.
        /// </summary>
        /// <param name="columnName">The name of the column (e.g., "A", "B", ...).</param>
        /// <param name="width">The desired width.</param>
        public void SetColumnWidth(string columnName, double width)
        {
            Columns columns = _worksheetPart.Worksheet.GetFirstChild<Columns>();
            if (columns == null)
            {
                columns = new Columns();
                _worksheetPart.Worksheet.InsertAfter(columns, _worksheetPart.Worksheet.GetFirstChild<SheetFormatProperties>());
            }

            uint columnIndex = (uint)ColumnLetterToIndex(columnName);
            var column = columns.Elements<Column>().FirstOrDefault(c => c.Min == columnIndex);
            if (column == null)
            {
                column = new Column { Min = columnIndex, Max = columnIndex, Width = width, CustomWidth = true };
                columns.Append(column);
            }
            else
            {
                column.Width = width;
                column.CustomWidth = true;
            }
        }

        /// <summary>
        /// Protects the worksheet with the specified password.
        /// </summary>
        /// <param name="password">The password to protect the worksheet.</param>
        public void ProtectSheet(string password)
        {
            SheetProtection sheetProtection = new SheetProtection()
            {
                Sheet = true,
                Objects = true,
                Scenarios = true,
                AutoFilter = true,
                PivotTables = true,
                Password = HashPassword(password),
                DeleteRows = true,
                DeleteColumns = true,
                FormatCells = true,
                FormatColumns = true,
                FormatRows = true,
                InsertColumns = true,
                InsertRows = true,
                InsertHyperlinks = true,
                Sort = true,
            };

            // Remove existing SheetProtection if any
            var existingProtection = _worksheetPart.Worksheet.Elements<SheetProtection>().FirstOrDefault();
            if (existingProtection != null)
            {
                existingProtection.Remove();
            }

            // Insert new SheetProtection after the SheetData element
            _worksheetPart.Worksheet.InsertAfter(sheetProtection, _worksheetPart.Worksheet.Elements<SheetData>().First());

            // Save the changes
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Checks if the worksheet is protected.
        /// </summary>
        /// <returns>True if the worksheet is protected, otherwise false.</returns>
        public bool IsProtected()
        {
            return _worksheetPart.Worksheet.Elements<SheetProtection>().Any();
        }


        /// <summary>
        /// Removes protection from the worksheet.
        /// </summary>
        public void UnprotectSheet()
        {
            if (IsProtected())
            {
                var sheetProtection = _worksheetPart.Worksheet.Elements<SheetProtection>().First();
                sheetProtection.Remove();
                
            }
            // Save the changes
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Hashes the given password.
        /// </summary>
        /// <param name="password">The password to hash.</param>
        /// <returns>The hashed password.</returns>
        private string HashPassword(string password)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(password));
        }



        /// <summary>
        /// Converts a column letter to its corresponding index.
        /// </summary>
        /// <param name="column">The column letter (e.g., "A", "B", ...).</param>
        /// <returns>The index corresponding to the column letter.</returns>
        private static int ColumnLetterToIndex(string column)
        {
            int index = 0;
            foreach (var ch in column)
            {
                index = (index * 26) + (ch - 'A' + 1);
            }
            return index;
        }


        /// <summary>
        /// Gets the file extension corresponding to a specific image content type.
        /// </summary>
        /// <param name="contentType">The image content type.</param>
        /// <returns>The file extension.</returns>
        private static string GetImageExtension(string contentType)
        {
            switch (contentType.ToLower())
            {
                case "image/jpeg": return "jpeg";
                case "image/png": return "png";
                case "image/gif": return "gif";
                case "image/tiff": return "tiff";
                case "image/bmp": return "bmp";
                default: throw new ArgumentOutOfRangeException(nameof(contentType), $"Unsupported image content type: {contentType}");
            }
        }


        /// <summary>
        /// Retrieves or creates a cell for a specific cell reference.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation.</param>
        /// <returns>The corresponding cell.</returns>
        private DocumentFormat.OpenXml.Spreadsheet.Cell GetOrCreateCell(string cellReference)
        {

            var cell = _sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                                .FirstOrDefault(c => string.Equals(c.CellReference.Value, cellReference, StringComparison.OrdinalIgnoreCase));

            if (cell == null)
            {
                cell = new DocumentFormat.OpenXml.Spreadsheet.Cell { CellReference = cellReference };
                var rowIndex = GetRowIndex(cellReference);
                var row = GetOrCreateRow(rowIndex);
                row.Append(cell);
            }

            return cell;
        }

        /// <summary>
        /// Retrieves the row index from a cell reference.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation.</param>
        /// <returns>The row index.</returns>
        private uint GetRowIndex(string cellReference)
        {
            var match = Regex.Match(cellReference, @"\d+");
            if (!match.Success)
                throw new FormatException("Invalid cell reference format.");

            return uint.Parse(match.Value);
        }

        /// <summary>
        /// Retrieves or creates a row for a specific row index.
        /// </summary>
        /// <param name="rowIndex">The row index.</param>
        /// <returns>The corresponding row.</returns>
        private Row GetOrCreateRow(uint rowIndex)
        {
            var row = _sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row { RowIndex = rowIndex };
                _sheetData.Append(row);
            }
            return row;
        }


        /// <summary>
        /// Merges cells within a specified range.
        /// </summary>
        /// <param name="startCellReference">The starting cell reference in A1 notation.</param>
        /// <param name="endCellReference">The ending cell reference in A1 notation.</param>
        public void MergeCells(string startCellReference, string endCellReference)
        {
            if (_worksheetPart.Worksheet.Elements<MergeCells>().Any())
            {
                // MergeCells element already exists, use it
                MergeCells mergeCells = _worksheetPart.Worksheet.Elements<MergeCells>().First();
                MergeCell newMergeCell = new MergeCell() { Reference = new StringValue(startCellReference + ":" + endCellReference) };
                mergeCells.Append(newMergeCell);
            }
            else
            {
                // Otherwise, create new MergeCells element
                MergeCells mergeCells = new MergeCells();
                MergeCell newMergeCell = new MergeCell() { Reference = new StringValue(startCellReference + ":" + endCellReference) };
                mergeCells.Append(newMergeCell);
                _worksheetPart.Worksheet.InsertAfter(mergeCells, _worksheetPart.Worksheet.Elements<SheetData>().First());
            }

            // Save changes to the WorksheetPart
            _worksheetPart.Worksheet.Save();
        }

        /// <summary>
        /// Retrieves the index of the worksheet within the workbook.
        /// </summary>
        /// <returns>The index of the worksheet.</returns>
        public int GetSheetIndex()
        {
            var workbookPart = _worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
            if (workbookPart == null)
                throw new InvalidOperationException("No WorkbookPart found.");

            var sheets = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
            var sheet = sheets.FirstOrDefault(s => workbookPart.GetPartById(s.Id) == _worksheetPart);

            if (sheet == null)
                throw new InvalidOperationException("Worksheet not found in workbook.");

            // Note: SheetId is not the same as the index of the sheet in the workbook.
            // If you specifically need the index, you may need to implement a different approach.
            return int.Parse(sheet.SheetId);
        }
    }

    public class CellIndexer
    {
        private readonly Worksheet _worksheet;

        /// <summary>
        /// Initializes a new instance of the <see cref="CellIndexer"/> class.
        /// </summary>
        /// <param name="worksheet">The worksheet to index.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="worksheet"/> is null.</exception>
        public CellIndexer(Worksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        /// <summary>
        /// Gets the cell at the specified reference.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation.</param>
        /// <returns>The cell at the specified reference.</returns>
        public Cell this[string cellReference]
        {
            get
            {
                // Delegate the actual work to Worksheet class
                return _worksheet.GetCell(cellReference);
            }
        }
    }
}

