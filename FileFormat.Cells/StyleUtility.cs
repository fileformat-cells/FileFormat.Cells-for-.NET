using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

public class StyleUtility
{
    private readonly WorkbookStylesPart stylesPart;
    private readonly Stylesheet stylesheet;

    public uint CreateDefaultStyle()
    {
        return CreateStyle("Arial", 12, "000000"); 
    }

    public StyleUtility(WorkbookStylesPart stylesPart)
    {
        this.stylesPart = stylesPart ?? throw new ArgumentNullException(nameof(stylesPart));
        this.stylesheet = this.stylesPart.Stylesheet ?? new Stylesheet();

        // Initialize Fonts, Fills, Borders, and CellFormats if they are null
        this.stylesheet.Fonts = this.stylesheet.Fonts ?? new Fonts();
        this.stylesheet.Fills = this.stylesheet.Fills ?? new Fills();
        this.stylesheet.Borders = this.stylesheet.Borders ?? new Borders();
        this.stylesheet.CellFormats = this.stylesheet.CellFormats ?? new CellFormats();
    }

    public uint CreateStyle(string fontName, double fontSize, string hexColor)
    {
        // Validate inputs
        if (string.IsNullOrEmpty(fontName))
            throw new ArgumentNullException(nameof(fontName));
        if (fontSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(fontSize), "Font size must be greater than zero");
        if (string.IsNullOrEmpty(hexColor) || !IsHexColor(hexColor))
            throw new ArgumentException("Invalid hex color", nameof(hexColor));

        // Create and append Font
        var font = new Font(
            new FontSize() { Val = fontSize },
            new Color() { Rgb = new HexBinaryValue() { Value = hexColor } },
            new FontName() { Val = fontName }
        );
        this.stylesheet.Fonts.Append(font);

        // Create and append Fill
        var fill = new Fill(new PatternFill() { PatternType = PatternValues.None });
        this.stylesheet.Fills.Append(fill);

        // Create and append Border
        var border = new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder());
        this.stylesheet.Borders.Append(border);

        // Create and append CellFormat
        var cellFormat = new CellFormat
        {
            FontId = new UInt32Value((uint)this.stylesheet.Fonts.ChildElements.Count - 1),
            FillId = new UInt32Value((uint)this.stylesheet.Fills.ChildElements.Count - 1),
            BorderId = new UInt32Value((uint)this.stylesheet.Borders.ChildElements.Count - 1)
        };
        this.stylesheet.CellFormats.Append(cellFormat);

        return (uint)this.stylesheet.CellFormats.ChildElements.Count - 1;
    }

    private static bool IsHexColor(string color)
    {
        return System.Text.RegularExpressions.Regex.IsMatch(color, "^(#)?([0-9a-fA-F]{3})([0-9a-fA-F]{3})?$");
    }


    public void SaveStylesheet()
    {
        if (this.stylesPart.Stylesheet == null)
        {
            this.stylesPart.Stylesheet = this.stylesheet;
        }
        this.stylesheet.Save();
    }
}
