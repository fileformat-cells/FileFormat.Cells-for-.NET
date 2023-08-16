namespace FileFormat.Cells
{
    /// <summary>
    /// Represents a row in the cell table.
    /// </summary>
    public class Row
    {
        /// <value>
        /// An object of the Parent Row class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Spreadsheet.Row row;

        /// <summary>
        /// Instantiate a new instance of the Row class.
        /// </summary>
        public Row()
        {
            this.row = new DocumentFormat.OpenXml.Spreadsheet.Row();
        }

        /// <summary>
        /// This property is used to set/get the Row Index.
        /// </summary>
        /// <returns>An integer value.</returns>
        public UInt32 Index {
            get { return this.row.RowIndex; }
            set {
                this.row.RowIndex = value;
            }
        }

        /// <summary>
        /// This method adds a Cell to a Row.
        /// </summary>
        /// <param name="cell">An object of the Cell class.</param>
        public void Append(Cell cell) {
            this.row.Append(cell.cell);
        }

    }
}

