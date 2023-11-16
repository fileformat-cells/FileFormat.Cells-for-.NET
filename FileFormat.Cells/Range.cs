using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileFormat.Cells
{
    public class Range
    {
        private readonly Worksheet _worksheet;

        public uint StartRowIndex { get; }
        public uint StartColumnIndex { get; }
        public uint EndRowIndex { get; }
        public uint EndColumnIndex { get; }

        // Returns the count of columns in the range
        public uint ColumnCount
        {
            get { return EndColumnIndex - StartColumnIndex + 1; }
        }

        // Returns the count of rows in the range
        public uint RowCount
        {
            get { return EndRowIndex - StartRowIndex + 1; }
        }

        
        public Range(Worksheet worksheet, uint startRowIndex, uint startColumnIndex, uint endRowIndex, uint endColumnIndex)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            StartRowIndex = startRowIndex;
            StartColumnIndex = startColumnIndex;
            EndRowIndex = endRowIndex;
            EndColumnIndex = endColumnIndex;
        }

        public void SetValue(string value)
        {
            for (uint row = StartRowIndex; row <= EndRowIndex; row++)
            {
                for (uint column = StartColumnIndex; column <= EndColumnIndex; column++)
                {
                    var cellReference = $"{ColumnIndexToLetter(column)}{row}";
                    var cell = _worksheet.GetCell(cellReference);
                    cell.PutValue(value);
                }
            }
        }

        public void ClearCells()
        {
            for (uint row = StartRowIndex; row <= EndRowIndex; row++)
            {
                for (uint column = StartColumnIndex; column <= EndColumnIndex; column++)
                {
                    var cellReference = $"{ColumnIndexToLetter(column)}{row}";
                    var cell = _worksheet.GetCell(cellReference);
                    if (cell != null)
                    {
                        cell.PutValue(string.Empty); // Clearing the value
                        cell.ApplyStyle(0); // Resetting the style if needed
                    }
                }
            }
        }

        public void MergeCells()
        {
            string startCellReference = $"{ColumnIndexToLetter(StartColumnIndex)}{StartRowIndex}";
            string endCellReference = $"{ColumnIndexToLetter(EndColumnIndex)}{EndRowIndex}";

            _worksheet.MergeCells(startCellReference, endCellReference);
        }


        public void AddDropdownListValidation(string[] options)
        {
            

            for (uint row = StartRowIndex; row <= EndRowIndex; row++)
            {
                for (uint column = StartColumnIndex; column <= EndColumnIndex; column++)
                {
                    var cellReference = $"{ColumnIndexToLetter(column)}{row}";
                    _worksheet.AddDropdownListValidation(cellReference, options);
                }
            }
        }



        private static string ColumnIndexToLetter(uint columnIndex)
        {
            string columnLetter = string.Empty;
            while (columnIndex > 0)
            {
                columnIndex--;
                columnLetter = (char)('A' + columnIndex % 26) + columnLetter;
                columnIndex /= 26;
            }
            return columnLetter;
        }
    }


}
