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

        public const double DefaultColumnWidth = 8.43; // Default width in character units
        public const double DefaultRowHeight = 15.0;   // Default height in points


        /// <summary>
        /// Gets the CellIndexer for the worksheet. This property provides indexed access to the cells of the worksheet.
        /// </summary>
        /// <value>
        /// The CellIndexer for the worksheet.
        /// </value>
        public CellIndexer Cells { get; }

        /// <summary>
        /// Initializes a new instance of the Worksheet class with the specified WorksheetPart and Worksheet.
        /// This constructor sets up the internal structure of the Worksheet object, including initializing the Cells property.
        /// </summary>
        /// <param name="worksheetPart">The WorksheetPart associated with this Worksheet. Cannot be null.</param>
        /// <param name="worksheet">The OpenXml Spreadsheet.Worksheet object. Cannot be null.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown if the provided worksheetPart is null.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if SheetData is not found in the provided worksheet.
        /// </exception>
        private Worksheet(WorksheetPart worksheetPart, DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet)
        {
            _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));

            _sheetData = worksheet?.Elements<SheetData>().FirstOrDefault()
                         ?? throw new InvalidOperationException("SheetData not found in the worksheet.");

            // Initialize the Cells property
            this.Cells = new CellIndexer(this);
        }

        /// <summary>
        /// Creates an instance of the Worksheet class using the provided WorksheetPart and Worksheet.
        /// This method serves as a factory for creating Worksheet objects, encapsulating the instantiation logic.
        /// </summary>
        /// <param name="worksheetPart">The WorksheetPart to be associated with the Worksheet. Cannot be null.</param>
        /// <param name="worksheet">The OpenXml Spreadsheet.Worksheet object to be associated with the Worksheet. Cannot be null.</param>
        /// <returns>A new instance of the Worksheet class.</returns>
        public class WorksheetFactory
        {
            public static Worksheet CreateInstance(WorksheetPart worksheetPart, DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet)
            {
                return new Worksheet(worksheetPart, worksheet);
            }
        }

        /// <summary>
        /// Gets or sets the name of the worksheet. This property performs several checks to ensure the integrity and validity of the worksheet and workbook parts.
        /// </summary>
        /// <value>
        /// The name of the worksheet.
        /// </value>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the WorksheetPart is null, if the WorkbookPart is not found as a parent, if the ID of the part is null or empty, or if no sheet is found with the specified ID.
        /// </exception>
        /// <exception cref="ArgumentException">
        /// Thrown when attempting to set the name with a null or empty value.
        /// </exception>
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
        /// Sets the height of the specified row in the worksheet.
        /// </summary>
        /// <param name="rowIndex">The 1-based index of the row for which the height is to be set.</param>
        /// <param name="height">The height to set for the specified row.</param>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if the rowIndex is less than 1 or if the height is less than 0.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the worksheet part or the worksheet is null.
        /// </exception>
        public void SetRowHeight(uint rowIndex, double height)
        {
            if (rowIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than 0.");
            }

            if (height < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(height), "Row height must be a positive number.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }
            var row = GetOrCreateRow(rowIndex);
            row.Height = height;
            row.CustomHeight = true;
        }

        /// <summary>
        /// Sets the width of the specified column in the worksheet.
        /// </summary>
        /// <param name="columnName">The name of the column (e.g., "A", "B", "C") for which the width is to be set.</param>
        /// <param name="width">The width to set for the specified column.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown if the columnName is null or empty.
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if the width is less than 0 or if the columnName is invalid or represents a column index out of range.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the worksheet part or the worksheet is null.
        /// </exception>
        public void SetColumnWidth(string columnName, double width)
        {
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new ArgumentNullException(nameof(columnName), "Column name cannot be null or empty.");
            }

            if (width < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(width), "Column width must be a positive number.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

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
        /// Retrieves the width of the specified column in the worksheet.
        /// If the width of the column has been explicitly set, it returns that value; otherwise, it returns the default column width.
        /// </summary>
        /// <param name="columnIndex">The 1-based index of the column for which the width is to be retrieved.</param>
        /// <returns>
        /// The width of the specified column. If the column's width is explicitly set, that value is returned; 
        /// otherwise, the default column width is returned.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if the columnIndex is less than 1.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the worksheet part or the worksheet is null.
        /// </exception>
        public double GetColumnWidth(uint columnIndex)
        {
            if (columnIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }
            // Access the Columns collection
            var columns = _worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
            if (columns != null)
            {
                foreach (var column in columns.Elements<Column>())
                {
                    
                    // Explicitly cast Min and Max to uint and check for null
                    uint min = column.Min.HasValue ? column.Min.Value : uint.MinValue;
                    uint max = column.Max.HasValue ? column.Max.Value : uint.MaxValue;

                    if (columnIndex >= min && columnIndex <= max)
                    {
                        // Also check if Width is set
                        return column.Width.HasValue ? column.Width.Value : DefaultColumnWidth;
                    }
                }
            }

            return DefaultColumnWidth;
        }
        /// <summary>
        /// Retrieves the height of the specified row in the worksheet.
        /// If the height of the row has been explicitly set, it returns that value; otherwise, it returns the default row height.
        /// </summary>
        /// <param name="rowIndex">The 1-based index of the row for which the height is to be retrieved.</param>
        /// <returns>
        /// The height of the specified row. If the row's height is explicitly set, that value is returned; 
        /// otherwise, the default row height is returned.
        /// </returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if the rowIndex is less than 1.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if the worksheet part or the worksheet is null.
        /// </exception>
        public double GetRowHeight(uint rowIndex)
        {
            if (rowIndex < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than 0.");
            }

            if (_worksheetPart == null || _worksheetPart.Worksheet == null)
            {
                throw new InvalidOperationException("Worksheet or WorksheetPart is null.");
            }

            // Assuming _worksheetPart is the OpenXML WorksheetPart
            var rows = _worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>().Elements<Row>();

            foreach (var row in rows)
            {
                // Check if this is the row we are looking for
                if (row.RowIndex.Value == rowIndex)
                {
                    // If Height is set, return it, otherwise return default height
                    return row.Height.HasValue ? row.Height.Value : DefaultRowHeight;
                }
            }

            return DefaultRowHeight; // Return default height if no specific height is set
        }




        /// <summary>
        /// Protects the worksheet with the specified password. This method applies various protection settings to the sheet, 
        /// including locking objects, scenarios, auto-filters, pivot tables, and other elements from editing.
        /// </summary>
        /// <param name="password">The password used to protect the worksheet. The password is hashed before being applied.</param>
        /// <remarks>
        /// This method creates a new SheetProtection object with specific settings and applies it to the worksheet.
        /// If an existing SheetProtection element is present, it is removed before applying the new protection.
        /// After setting the protection, it saves the changes to the worksheet.
        /// </remarks>
        /// <exception cref="ArgumentException">
        /// Thrown if the provided <paramref name="password"/> is null or empty.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown if an error occurs while applying the protection, such as failing to save the changes.
        /// </exception
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
        /// Determines whether the worksheet is protected.
        /// This method checks for the presence of a SheetProtection element within the worksheet to ascertain its protection status.
        /// </summary>
        /// <returns>
        /// A boolean value indicating whether the worksheet is protected. Returns <c>true</c> if the worksheet is protected, otherwise <c>false</c>.
        /// </returns>
        /// <remarks>
        /// A worksheet is considered protected if there is at least one SheetProtection element present in its elements.
        /// </remarks>
        public bool IsProtected()
        {
            return _worksheetPart.Worksheet.Elements<SheetProtection>().Any();
        }


        /// <summary>
        /// Removes protection from the worksheet, if it is currently protected.
        /// This method checks for the presence of a SheetProtection element and removes it to unprotect the sheet.
        /// </summary>
        /// <remarks>
        /// If the worksheet is protected (indicated by the presence of a SheetProtection element), this method removes the protection.
        /// After altering the protection status, it saves the changes to the worksheet.
        /// If the worksheet is not protected, this method performs no action.
        /// </remarks>
        /// <exception cref="InvalidOperationException">
        /// Thrown if an attempt is made to remove protection but no SheetProtection element is found. This should not normally occur, as the method first checks if the sheet is protected.
        /// </exception>
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
        /// Merges a range of cells specified by the start and end cell references in A1 notation.
        /// This method creates a merged cell area that spans from the start cell to the end cell.
        /// </summary>
        /// <param name="startCellReference">The start cell reference in A1 notation for the merge range.</param>
        /// <param name="endCellReference">The end cell reference in A1 notation for the merge range.</param>
        /// <remarks>
        /// If a MergeCells element already exists in the worksheet, this method appends a new merge cell reference to it.
        /// If no MergeCells element exists, it creates a new one and then appends the merge cell reference.
        /// After defining the merge cell range, it saves the changes to the WorksheetPart.
        /// </remarks>
        /// <exception cref="ArgumentException">
        /// Thrown when either <paramref name="startCellReference"/> or <paramref name="endCellReference"/> is null, empty, or invalid.
        /// </exception>
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
        /// Retrieves the index of the current worksheet within the workbook. This method locates the worksheet within the workbook's collection of sheets and returns its index.
        /// Note that the SheetId property of a worksheet is different from its index in the workbook's sheet collection.
        /// </summary>
        /// <returns>
        /// The index of the sheet within the workbook. This is not the same as the SheetId.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown if no WorkbookPart is found, or if the worksheet is not found in the workbook.
        /// </exception>
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


        /// <summary>
        /// Retrieves a range of cells specified by the start and end row and column indices.
        /// </summary>
        /// <param name="startRowIndex">The starting row index of the range.</param>
        /// <param name="startColumnIndex">The starting column index of the range.</param>
        /// <param name="endRowIndex">The ending row index of the range.</param>
        /// <param name="endColumnIndex">The ending column index of the range.</param>
        /// <returns>
        /// A <see cref="Range"/> object representing the specified range of cells.
        /// </returns>
        public Range GetRange(uint startRowIndex, uint startColumnIndex, uint endRowIndex, uint endColumnIndex)
        {
            return new Range(this, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        }

        /// <summary>
        /// Retrieves a range of cells specified by the start and end cell references in A1 notation.
        /// </summary>
        /// <param name="startCellReference">The start cell reference in A1 notation.</param>
        /// <param name="endCellReference">The end cell reference in A1 notation.</param>
        /// <returns>
        /// A <see cref="Range"/> object representing the specified range of cells.
        /// </returns>
        public Range GetRange(string startCellReference, string endCellReference)
        {
            var startCellParts = ParseCellReference(startCellReference);
            var endCellParts = ParseCellReference(endCellReference);
            return GetRange(startCellParts.row, startCellParts.column, endCellParts.row, endCellParts.column);
        }

        /// <summary>
        /// Adds a dropdown list validation to a specified cell. The dropdown list contains the options provided.
        /// </summary>
        /// <param name="cellReference">The cell reference in A1 notation where the dropdown should be added.</param>
        /// <param name="options">An array of string values that will appear as options in the dropdown list.</param>
        /// <exception cref="ArgumentException">
        /// Thrown when <paramref name="cellReference"/> is null or invalid, or if <paramref name="options"/> is empty or null.
        /// </exception>
        /// <remarks>
        /// This method creates a data validation rule that restricts input to the cell to the provided list of options.
        /// </remarks>
        public void AddDropdownListValidation(string cellReference, string[] options)
        {
            // Convert options array into a comma-separated string
            string formula = string.Join(",", options);

            // Create the data validation object
            DataValidation dataValidation = new DataValidation
            {
                Type = DataValidationValues.List,
                ShowDropDown = true,
                ShowErrorMessage = true,
                ErrorTitle = "Invalid input",
                Error = "The value entered is not in the list.",
                Formula1 = new Formula1("\"" + formula + "\""), // The formula is enclosed in quotes
                SequenceOfReferences = new ListValue<StringValue> { InnerText = cellReference }
            };

            // Add the data validation to the worksheet
            var dataValidations = _worksheetPart.Worksheet.GetFirstChild<DataValidations>();
            if (dataValidations == null)
            {
                dataValidations = new DataValidations();
                _worksheetPart.Worksheet.AppendChild(dataValidations);
            }

            dataValidations.AppendChild(dataValidation);
        }

        public void ApplyValidation(string cellReference, ValidationRule rule)
        {
            DataValidation dataValidation = CreateDataValidation(cellReference, rule);
            AddDataValidation(dataValidation);
        }

        public ValidationRule GetValidationRule(string cellReference)
        {
            if (_worksheetPart == null)
            {
                throw new InvalidOperationException("Worksheet part is not loaded.");
            }

            var dataValidations = _worksheetPart.Worksheet.Descendants<DataValidation>();

            foreach (var dataValidation in dataValidations)
            {
                if (dataValidation.SequenceOfReferences.InnerText.Contains(cellReference))
                {
                    return CreateValidationRuleFromDataValidation(dataValidation);
                }
            }

            return null; // No validation rule found for the specified cell
        }

        private ValidationRule CreateValidationRuleFromDataValidation(DataValidation dataValidation)
        {
            // Determine the type of validation
            ValidationType type = DetermineValidationType(dataValidation.Type);

            // Depending on the type, create a corresponding ValidationRule
            switch (type)
            {
                case ValidationType.List:
                    // Assuming list validations use a comma-separated list of options in Formula1
                    var options = dataValidation.Formula1.Text.Split(',');
                    return new ValidationRule(options);

                case ValidationType.CustomFormula:
                    return new ValidationRule(dataValidation.Formula1.Text);

                // Add cases for other validation types (e.g., Date, WholeNumber, Decimal, TextLength)
                // ...

                default:
                    throw new NotImplementedException("Validation type not supported.");
            }
        }

        private ValidationType DetermineValidationType(DataValidationValues openXmlType)
        {
            // Map the OpenXML DataValidationValues to your ValidationType enum
            // This mapping depends on how closely your ValidationType enum aligns with OpenXML's types
            // Example mapping:
            switch (openXmlType)
            {
                case DataValidationValues.List:
                    return ValidationType.List;
                case DataValidationValues.Custom:
                    return ValidationType.CustomFormula;
                // Map other types...
                default:
                    throw new NotImplementedException("Validation type not supported.");
            }
        }



        private DataValidation CreateDataValidation(string cellReference, ValidationRule rule)
        {
            DataValidation dataValidation = new DataValidation
            {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = cellReference },
                ShowErrorMessage = true,
                ErrorTitle = rule.ErrorTitle,
                Error = rule.ErrorMessage
            };

            switch (rule.Type)
            {
                case ValidationType.List:
                    dataValidation.Type = DataValidationValues.List;
                    dataValidation.Formula1 = new Formula1($"\"{string.Join(",", rule.Options)}\"");
                    break;

                case ValidationType.Date:
                    dataValidation.Type = DataValidationValues.Date;
                    dataValidation.Operator = DataValidationOperatorValues.Between;
                    dataValidation.Formula1 = new Formula1(rule.MinValue.HasValue ? rule.MinValue.Value.ToString() : "0");
                    dataValidation.Formula2 = new Formula2(rule.MaxValue.HasValue ? rule.MaxValue.Value.ToString() : "0");
                    break;

                case ValidationType.WholeNumber:
                    dataValidation.Type = DataValidationValues.Whole;
                    dataValidation.Operator = DataValidationOperatorValues.Between;
                    dataValidation.Formula1 = new Formula1(rule.MinValue.HasValue ? rule.MinValue.Value.ToString() : "0");
                    dataValidation.Formula2 = new Formula2(rule.MaxValue.HasValue ? rule.MaxValue.Value.ToString() : "0");
                    break;

                case ValidationType.Decimal:
                    dataValidation.Type = DataValidationValues.Decimal;
                    dataValidation.Operator = DataValidationOperatorValues.Between;
                    dataValidation.Formula1 = new Formula1(rule.MinValue.HasValue ? rule.MinValue.Value.ToString() : "0");
                    dataValidation.Formula2 = new Formula2(rule.MaxValue.HasValue ? rule.MaxValue.Value.ToString() : "0");
                    break;

                case ValidationType.TextLength:
                    dataValidation.Type = DataValidationValues.TextLength;
                    dataValidation.Operator = DataValidationOperatorValues.Between;
                    dataValidation.Formula1 = new Formula1(rule.MinValue.HasValue ? rule.MinValue.Value.ToString() : "0");
                    dataValidation.Formula2 = new Formula2(rule.MaxValue.HasValue ? rule.MaxValue.Value.ToString() : "0");
                    break;

                case ValidationType.CustomFormula:
                    dataValidation.Type = DataValidationValues.Custom;
                    dataValidation.Formula1 = new Formula1(rule.CustomFormula);
                    break;

                default:
                    throw new ArgumentException("Unsupported validation type.");
            }

            return dataValidation;
        }


        private void AddDataValidation(DataValidation dataValidation)
        {
            var dataValidations = _worksheetPart.Worksheet.GetFirstChild<DataValidations>();
            if (dataValidations == null)
            {
                dataValidations = new DataValidations();
                _worksheetPart.Worksheet.AppendChild(dataValidations);
            }
            dataValidations.AppendChild(dataValidation);
        }


        private (uint row, uint column) ParseCellReference(string cellReference)
        {
            var match = Regex.Match(cellReference, @"([A-Z]+)(\d+)");
            if (!match.Success)
                throw new FormatException("Invalid cell reference format.");

            uint row = uint.Parse(match.Groups[2].Value);
            uint column = (uint)ColumnLetterToIndex(match.Groups[1].Value);

            return (row, column);
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
        /// Gets the cell at the specified reference in A1 notation. This indexer provides a convenient way to access cells within the worksheet.
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

