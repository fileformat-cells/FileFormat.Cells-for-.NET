namespace FileFormat.Cells
{
    /// <summary>
    /// Represents the cell table, grouped together by rows.
    /// </summary>
    public class SheetData
    {
        /// <value>
        /// An object of the Parent SheetData class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Spreadsheet.SheetData sheetData;

        /// <summary>
        /// Instantiate a new instance of the SheetData class.
        /// </summary>
        public SheetData()
        {
            this.sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
        }

        /// <summary>
        /// This method adds a Row to SheetData.
        /// </summary>
        /// <param name="row">An object of the Row class.</param>
        public void Append(Row row) {
            this.sheetData.Append(row.row);
        }

        /// <summary>
        /// Gets the underlying SheetData object.
        /// </summary>
        internal DocumentFormat.OpenXml.Spreadsheet.SheetData GetSheetData()
        {
            return this.sheetData;
        }

    }
}

