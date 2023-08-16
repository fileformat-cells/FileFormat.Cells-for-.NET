namespace FileFormat.Cells
{
    /// <summary>
    /// Represents the style information for a cell within a spreadsheet.
    /// </summary>
	public class CellStyle
	{
        /// <summary>
        /// Gets or sets the font size for the cell's content.
        /// </summary>
        /// <value>The size of the font.</value>
        public int? FontSize { get; set; }

        /// <summary>
        /// Gets or sets the font name for the cell's content.
        /// </summary>
        /// <value>The name of the font.</value>
        public string? FontName { get; set; }

        /// <summary>
        /// Gets or sets the background color of the cell.
        /// </summary>
        /// <value>The color of the cell in a hex format (e.g., "#FF0000" for red).</value>
        public string? CellColor { get; set; }
    }
}

