using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace FileFormat.Cells.Image
{
    /// <summary>
    /// This class contains methods to add images to an Excel document.
    /// </summary>
    public class Image
    {
        /// <value>
        /// An object of the Parent WorksheetPart class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart;

        /// <value>
        /// An object of the Parent Drawing class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Spreadsheet.Drawing drawing;

        /// <value>
        /// An object of the Parent SpreadsheetDocument class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheetDocument;
        private Xdr.WorksheetDrawing worksheetDrawing1;
        private int imageCount = 0;
        private int currentSheet = 0;

        /// <summary>
        /// Instantiates a new instance of the Image class.
        /// </summary>
        /// <param name="workbook">An object of the Workbook class.</param>
        public Image(Workbook workbook)
        {
            this.spreadsheetDocument = workbook.spreadsheetDocument;
            worksheetDrawing1 = new Xdr.WorksheetDrawing();

        }
        /// <summary>
        /// Invoke this function to add images into a Worksheet.
        /// </summary>
        /// <param name="sheetIndex">An integer value represents the index of worksheet.</param>
        ///  <param name="imagePath">An integer value represents the index of worksheet.</param>
        ///  <param name="startRowIndex">An integer value that represents the starting row index.</param>
        ///  <param name="startColumnIndex">An integer value that represents the starting column index.</param>
        ///  <param name="endRowIndex">An integer value that represents the ending row index.</param>
        ///  <param name="endColumnIndex">An integer value that represents the ending column index.</param>
        
        public void Add(int sheetIndex, string imagePath, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex) {
 
            DrawingsPart drawingsPart;
            imageCount++;
            this.drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = "rId"+ imageCount.ToString() };

            if (this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].DrawingsPart == null)
            {
                drawingsPart = this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].AddNewPart<DrawingsPart>("rId" + imageCount.ToString());
                this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].Worksheet.Append(drawing);
            }
            else 
                drawingsPart = this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].DrawingsPart;
            
            GenerateDrawingsPart(sheetIndex,drawingsPart, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
            ImagePart imagePart1 = drawingsPart.AddNewPart<ImagePart>("image/jpeg", "rId" + imageCount);
            Stream imageStream = File.Open(imagePath, FileMode.Open);
            imagePart1.FeedData(imageStream);
            
        }


        private void GenerateDrawingsPart(int sheetIndex, DrawingsPart drawingsPart1, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
           
            if (imageCount == 1 && sheetIndex == currentSheet)
            {
                currentSheet = sheetIndex;

                worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

                Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
                Xdr.ColumnId columnId1 = new Xdr.ColumnId();
                columnId1.Text = startColumnIndex.ToString();
                Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
                columnOffset1.Text = "38100";
                Xdr.RowId rowId1 = new Xdr.RowId();
                rowId1.Text = startRowIndex.ToString();
                Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
                rowOffset1.Text = "0";

                fromMarker1.Append(columnId1);
                fromMarker1.Append(columnOffset1);
                fromMarker1.Append(rowId1);
                fromMarker1.Append(rowOffset1);

                Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
                Xdr.ColumnId columnId2 = new Xdr.ColumnId();
                columnId2.Text = endColumnIndex.ToString();
                Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
                columnOffset2.Text = "542925";
                Xdr.RowId rowId2 = new Xdr.RowId();
                rowId2.Text = endRowIndex.ToString();
                Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
                rowOffset2.Text = "161925";

                toMarker1.Append(columnId2);
                toMarker1.Append(columnOffset2);
                toMarker1.Append(rowId2);
                toMarker1.Append(rowOffset2);

                Xdr.Picture picture1 = new Xdr.Picture();

                Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
                Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

                Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
                A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

                nonVisualPictureDrawingProperties1.Append(pictureLocks1);

                nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
                nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

                Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

                A.Blip blip1 = new A.Blip() { Embed = "rId" + imageCount, CompressionState = A.BlipCompressionValues.Print };
                blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

                A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

                A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
                useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                blipExtension1.Append(useLocalDpi1);

                blipExtensionList1.Append(blipExtension1);

                blip1.Append(blipExtensionList1);

                A.Stretch stretch1 = new A.Stretch();
                A.FillRectangle fillRectangle1 = new A.FillRectangle();

                stretch1.Append(fillRectangle1);

                blipFill1.Append(blip1);
                blipFill1.Append(stretch1);

                Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset1 = new A.Offset() { X = 1257300L, Y = 762000L };
                A.Extents extents1 = new A.Extents() { Cx = 2943225L, Cy = 2257425L };

                transform2D1.Append(offset1);
                transform2D1.Append(extents1);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);

                picture1.Append(nonVisualPictureProperties1);
                picture1.Append(blipFill1);
                picture1.Append(shapeProperties1);
                Xdr.ClientData clientData1 = new Xdr.ClientData();

                twoCellAnchor1.Append(fromMarker1);
                twoCellAnchor1.Append(toMarker1);
                twoCellAnchor1.Append(picture1);
                twoCellAnchor1.Append(clientData1);

                worksheetDrawing1.Append(twoCellAnchor1);
                if (imageCount == 1)
                    drawingsPart1.WorksheetDrawing = worksheetDrawing1;
            }
            else if (imageCount == 1 && sheetIndex != currentSheet)
            {
                currentSheet = sheetIndex;

                worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

                Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
                Xdr.ColumnId columnId1 = new Xdr.ColumnId();
                columnId1.Text = startColumnIndex.ToString();
                Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
                columnOffset1.Text = "38100";
                Xdr.RowId rowId1 = new Xdr.RowId();
                rowId1.Text = startRowIndex.ToString();
                Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
                rowOffset1.Text = "0";

                fromMarker1.Append(columnId1);
                fromMarker1.Append(columnOffset1);
                fromMarker1.Append(rowId1);
                fromMarker1.Append(rowOffset1);

                Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
                Xdr.ColumnId columnId2 = new Xdr.ColumnId();
                columnId2.Text = endColumnIndex.ToString();
                Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
                columnOffset2.Text = "542925";
                Xdr.RowId rowId2 = new Xdr.RowId();
                rowId2.Text = endRowIndex.ToString();
                Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
                rowOffset2.Text = "161925";

                toMarker1.Append(columnId2);
                toMarker1.Append(columnOffset2);
                toMarker1.Append(rowId2);
                toMarker1.Append(rowOffset2);

                Xdr.Picture picture1 = new Xdr.Picture();

                Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
                Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

                Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
                A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

                nonVisualPictureDrawingProperties1.Append(pictureLocks1);

                nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
                nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

                Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

                A.Blip blip1 = new A.Blip() { Embed = "rId" + imageCount, CompressionState = A.BlipCompressionValues.Print };
                blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

                A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

                A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
                useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                blipExtension1.Append(useLocalDpi1);

                blipExtensionList1.Append(blipExtension1);

                blip1.Append(blipExtensionList1);

                A.Stretch stretch1 = new A.Stretch();
                A.FillRectangle fillRectangle1 = new A.FillRectangle();

                stretch1.Append(fillRectangle1);

                blipFill1.Append(blip1);
                blipFill1.Append(stretch1);

                Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset1 = new A.Offset() { X = 1257300L, Y = 762000L };
                A.Extents extents1 = new A.Extents() { Cx = 2943225L, Cy = 2257425L };

                transform2D1.Append(offset1);
                transform2D1.Append(extents1);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);

                picture1.Append(nonVisualPictureProperties1);
                picture1.Append(blipFill1);
                picture1.Append(shapeProperties1);
                Xdr.ClientData clientData1 = new Xdr.ClientData();

                twoCellAnchor1.Append(fromMarker1);
                twoCellAnchor1.Append(toMarker1);
                twoCellAnchor1.Append(picture1);
                twoCellAnchor1.Append(clientData1);

                worksheetDrawing1.Append(twoCellAnchor1);
                if (imageCount == 1)
                    drawingsPart1.WorksheetDrawing = worksheetDrawing1;
            }
            else if (imageCount > 1 && sheetIndex == currentSheet) {
                currentSheet = sheetIndex;

                Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

                Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
                Xdr.ColumnId columnId1 = new Xdr.ColumnId();
                columnId1.Text = startColumnIndex.ToString();
                Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
                columnOffset1.Text = "38100";
                Xdr.RowId rowId1 = new Xdr.RowId();
                rowId1.Text = startRowIndex.ToString();
                Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
                rowOffset1.Text = "0";

                fromMarker1.Append(columnId1);
                fromMarker1.Append(columnOffset1);
                fromMarker1.Append(rowId1);
                fromMarker1.Append(rowOffset1);

                Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
                Xdr.ColumnId columnId2 = new Xdr.ColumnId();
                columnId2.Text = endColumnIndex.ToString();
                Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
                columnOffset2.Text = "542925";
                Xdr.RowId rowId2 = new Xdr.RowId();
                rowId2.Text = endRowIndex.ToString();
                Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
                rowOffset2.Text = "161925";

                toMarker1.Append(columnId2);
                toMarker1.Append(columnOffset2);
                toMarker1.Append(rowId2);
                toMarker1.Append(rowOffset2);

                Xdr.Picture picture1 = new Xdr.Picture();

                Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
                Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

                Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
                A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

                nonVisualPictureDrawingProperties1.Append(pictureLocks1);

                nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
                nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

                Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

                A.Blip blip1 = new A.Blip() { Embed = "rId" + imageCount, CompressionState = A.BlipCompressionValues.Print };
                blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

                A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

                A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
                useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                blipExtension1.Append(useLocalDpi1);

                blipExtensionList1.Append(blipExtension1);

                blip1.Append(blipExtensionList1);

                A.Stretch stretch1 = new A.Stretch();
                A.FillRectangle fillRectangle1 = new A.FillRectangle();

                stretch1.Append(fillRectangle1);

                blipFill1.Append(blip1);
                blipFill1.Append(stretch1);

                Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset1 = new A.Offset() { X = 1257300L, Y = 762000L };
                A.Extents extents1 = new A.Extents() { Cx = 2943225L, Cy = 2257425L };

                transform2D1.Append(offset1);
                transform2D1.Append(extents1);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);

                picture1.Append(nonVisualPictureProperties1);
                picture1.Append(blipFill1);
                picture1.Append(shapeProperties1);
                Xdr.ClientData clientData1 = new Xdr.ClientData();

                twoCellAnchor1.Append(fromMarker1);
                twoCellAnchor1.Append(toMarker1);
                twoCellAnchor1.Append(picture1);
                twoCellAnchor1.Append(clientData1);

                worksheetDrawing1.Append(twoCellAnchor1);
            }
        }

        /// <summary>
        /// Invoke this method to read the value of a particular cell. 
        /// </summary>
        /// <param name="sheetIndex">An Integer value.</param>
        /// <returns>A collection of image streams.</returns>
        public List<Stream> ExtractImagesFromWorkSheet(int sheetIndex)
        {
            List<Stream> imagePartsCollection = new List<Stream>();
            if (this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].Worksheet.WorksheetPart.DrawingsPart == null)
                return null;
            foreach (var part in this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[sheetIndex].Worksheet.WorksheetPart.DrawingsPart.ImageParts)
            {
                imagePartsCollection.Add(part.GetStream());
            }
            return imagePartsCollection;
        }

        /// <summary>
        /// Call this property to get the total number of images in a Workbook.
        /// </summary>
        /// <returns>The count of the total number of images in a Workbook. </returns>
        public int GetImagesCountFromWorkBook
        {
            get
            {
                int count = 0;
                for (int index = 0; index < this.spreadsheetDocument.WorkbookPart.WorksheetParts.Count(); index++)
                {
                    if (this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[index].Worksheet.WorksheetPart.DrawingsPart == null)
                        continue;
                    foreach (var part in this.spreadsheetDocument.WorkbookPart.WorksheetParts.ToList()[index].Worksheet.WorksheetPart.DrawingsPart.ImageParts)
                    {
                        count++;
                    }
                }
                return count;
            }
        }

    }
}

