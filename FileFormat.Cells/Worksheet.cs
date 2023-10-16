using System;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FileFormat.Cells
{
    public sealed class Worksheet
    {
        private WorksheetPart _worksheetPart;
        private SheetData _sheetData;

        // New Cells property
        public CellIndexer Cells { get; }

        internal Worksheet(WorksheetPart worksheetPart, DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet)
        {
            _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));

            _sheetData = worksheet?.Elements<SheetData>().FirstOrDefault()
                         ?? throw new InvalidOperationException("SheetData not found in the worksheet.");

            // Initialize the Cells property
            this.Cells = new CellIndexer(this);
        }

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



        // New GetCell method
        public Cell GetCell(string cellReference)
        {
            // This logic used to be in your indexer
            return new Cell(GetOrCreateCell(cellReference), _sheetData);
        }

        public void AddImage(Image image, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            if (image == null) throw new ArgumentNullException(nameof(image));

            // Assuming you have a working constructor or factory method for ImageHandler
            var imgHandler = new ImageHandler(_worksheetPart);
            imgHandler.Add(image.Path, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        }


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

        public void SetRowHeight(uint rowIndex, double height)
        {
            var row = GetOrCreateRow(rowIndex);
            row.Height = height;
            row.CustomHeight = true;
        }

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

        public bool IsProtected()
        {
            return _worksheetPart.Worksheet.Elements<SheetProtection>().Any();
        }

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


        private string HashPassword(string password)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(password));
        }




        private static int ColumnLetterToIndex(string column)
        {
            int index = 0;
            foreach (var ch in column)
            {
                index = (index * 26) + (ch - 'A' + 1);
            }
            return index;
        }





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

        private uint GetRowIndex(string cellReference)
        {
            var match = Regex.Match(cellReference, @"\d+");
            if (!match.Success)
                throw new FormatException("Invalid cell reference format.");

            return uint.Parse(match.Value);
        }

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

        public CellIndexer(Worksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

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

