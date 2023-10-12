using System;
using DocumentFormat.OpenXml;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FileFormat.Cells
{
    public sealed class Cell
    {
        private readonly DocumentFormat.OpenXml.Spreadsheet.Cell _cell;
        private readonly SheetData _sheetData;

        public string CellReference => _cell.CellReference;

        public Cell(DocumentFormat.OpenXml.Spreadsheet.Cell cell, SheetData sheetData)
        {
            _cell = cell ?? throw new ArgumentNullException(nameof(cell));
            _sheetData = sheetData ?? throw new ArgumentNullException(nameof(sheetData));
        }

        public void PutValue(string value)
        {
            PutValue(value, CellValues.String);
        }

        public void PutValue(double value)
        {
            PutValue(value.ToString(CultureInfo.InvariantCulture), CellValues.Number);
        }

        public void PutValue(DateTime value)
        {
            PutValue(value.ToOADate().ToString(CultureInfo.InvariantCulture), CellValues.Date);
        }

        private void PutValue(string value, CellValues dataType)
        {
            _cell.DataType = new EnumValue<CellValues>(dataType);
            _cell.CellValue = new CellValue(value);

        }

        public void PutFormula(string formula)
        {
            _cell.CellFormula = new CellFormula(formula);
            _cell.CellValue = new CellValue(); // You might want to set some default value or calculated value here
        }

        public string GetValue()
        {
            return _cell.CellValue?.Text;
        }

        public CellValues? GetDataType()
        {
            return _cell.DataType?.Value;
        }

        

        public string GetFormula()
        {
            return _cell.CellFormula?.Text;
        }

        public void ApplyStyle(uint styleIndex)
        {
            _cell.StyleIndex = styleIndex;
        }
    }

}

