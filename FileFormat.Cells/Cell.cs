using DocumentFormat.OpenXml;

namespace FileFormat.Cells
{
    /// <summary>
    /// Represents a cell in a row.
    /// </summary>
    public class Cell
    {
        /// <value>
        /// An object of the Parent Cell class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Spreadsheet.Cell cell;
        private Styles styles;

        /// <summary>
        /// Instantiate a new instance of the Cell class.
        /// </summary>
        public Cell()
        {
            this.cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
            //this.styles = new Styles();
        }

        /// <summary>
        /// This method is used to set the Cell Reference in a worksheet. 
        /// </summary>
        /// <param name="value">A string value.</param>
        public void setCellReference(string value)
        {
            this.cell.CellReference = value;
        }

        /// <summary>
        /// This method is used to set the Cell data type to String.
        /// </summary>
        public void setStringDataType()
        {
            this.cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
        }
        /// <summary>
        /// This method is used to set the Cell data type to Number.
        /// </summary>
        public void setNumberDataType()
        {
            this.cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
        }

        /// <summary>
        /// This method is used to set the value of a Cell.
        /// </summary>
        /// <param name="value">A dynamic value.</param>
        public void CellValue(dynamic value)
        {
            this.cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value);
        }

        /// <summary>
        /// Sets the style index of the cell to 1.
        /// </summary>
        public void CellIndex()
        {
            this.cell.StyleIndex = 1;
        }

        /// <summary>
        /// Sets the style index of the cell to the specified value.
        /// </summary>
        /// <param name="num">The style index is to be set for the cell.</param>

        public void CellIndex(UInt32Value num)
        {
            this.cell.StyleIndex = num;
        }


        // Other properties and methods...
    }

}

