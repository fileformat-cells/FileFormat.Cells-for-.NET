using System;
using DocumentFormat.OpenXml;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace FileFormat.Cells
{
    public sealed class Cell
    {

        private readonly DocumentFormat.OpenXml.Spreadsheet.Cell _cell;
        private readonly WorkbookPart _workbookPart;

        private readonly SheetData _sheetData;

        /// <summary>
        /// Gets the cell reference in A1 notation.
        /// </summary>
        public string CellReference => _cell.CellReference;

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class.
        /// </summary>
        /// <param name="cell">The underlying OpenXML cell object.</param>
        /// <param name="sheetData">The sheet data containing the cell.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="cell"/> or <paramref name="sheetData"/> is null.
        /// </exception>
        public Cell(DocumentFormat.OpenXml.Spreadsheet.Cell cell, SheetData sheetData, WorkbookPart workbookPart)
        {
            _cell = cell ?? throw new ArgumentNullException(nameof(cell));
            _sheetData = sheetData ?? throw new ArgumentNullException(nameof(sheetData));
            _workbookPart = workbookPart ?? throw new ArgumentNullException(nameof(workbookPart));
        }

        /// <summary>
        /// Sets the value of the cell as a string.
        /// </summary>
        /// <param name="value">The value to set.</param>
        public void PutValue(string value)
        {
            PutValue(value, CellValues.String);
        }

        /// <summary>
        /// Sets the value of the cell as a number.
        /// </summary>
        /// <param name="value">The numeric value to set.</param>
        public void PutValue(double value)
        {
            PutValue(value.ToString(CultureInfo.InvariantCulture), CellValues.Number);
        }

        /// <summary>
        /// Sets the value of the cell as a date.
        /// </summary>
        /// <param name="value">The date value to set.</param>
        public void PutValue(DateTime value)
        {
            PutValue(value.ToOADate().ToString(CultureInfo.InvariantCulture), CellValues.Date);
        }

        /// <summary>
        /// Sets the cell's value with a specific data type.
        /// </summary>
        /// <param name="value">The value to set.</param>
        /// <param name="dataType">The data type of the value.</param>
        private void PutValue(string value, CellValues dataType)
        {
            _cell.DataType = new EnumValue<CellValues>(dataType);
            _cell.CellValue = new CellValue(value);

        }

        /// <summary>
        /// Sets a formula for the cell.
        /// </summary>
        /// <param name="formula">The formula to set.</param>
        public void PutFormula(string formula)
        {
            _cell.CellFormula = new CellFormula(formula);
            _cell.CellValue = new CellValue(); // You might want to set some default value or calculated value here
        }

        /// <summary>
        /// Gets the value of the cell.
        /// </summary>
        /// <returns>The cell value as a string.</returns>
        public string GetValue()
        {
            if (_cell == null || _cell.CellValue == null) return "";

            if (_cell.DataType != null && _cell.DataType.Value == CellValues.SharedString)
            {
                int index = int.Parse(_cell.CellValue.Text);
                SharedStringTablePart sharedStrings = _workbookPart.SharedStringTablePart;
                return sharedStrings.SharedStringTable.ElementAt(index).InnerText;
            }
            else
            {
                return _cell.CellValue.Text;
            }
        }

        /// <summary>
        /// Gets the data type of the cell's value.
        /// </summary>
        /// <returns>The cell's value data type, or null if not set.</returns>
        public CellValues? GetDataType()
        {
            return _cell.DataType?.Value;
        }


        /// <summary>
        /// Gets the formula set for the cell.
        /// </summary>
        /// <returns>The cell's formula as a string, or null if not set.</returns>
        public string GetFormula()
        {
            return _cell.CellFormula?.Text;
        }

        /// <summary>
        /// Applies a style to the cell.
        /// </summary>
        /// <param name="styleIndex">The index of the style to apply.</param>
        public void ApplyStyle(uint styleIndex)
        {
            _cell.StyleIndex = styleIndex;
        }
    }

}

