using System.ComponentModel;
using FileFormat.Cells;
using FileFormat.Cells.Image;

class Program
{
    static void Main(string[] args)
    {
        // for creating a worksbook with some data 
        string filePath = "/Users/Mustafa/Desktop/Cells/FileFormat.Cells/TestSpreadSheets/test.xlsx";
        string imagePath = "/Users/Mustafa/Desktop/Cells/FileFormat.Cells/TestSpreadSheets/pic.png";
        string imagePath2 = "/Users/Mustafa/Desktop/Cells/FileFormat.Cells/TestSpreadSheets/pic2.png";

        // for creating an empty Excel file 
        //Workbook workbook = new Workbook();
        //workbook.Save(filePath);

        //*****************************************************************

        // for creating an Excel file with worksheets containg some data  
        //Workbook workbook = new Workbook();
        //Worksheet worksheet = new Worksheet(workbook);

        //CellStyle cellStyle = new CellStyle()
        //{
        //    CellColor = "18FF6D",
        //    FontName = "Calibri",
        //    FontSize = 22
        //};
        //uint format_cell_id = worksheet.insertStyle(cellStyle);

        //worksheet.insertValue("A1", 1, " some data ", format_cell_id);
        //worksheet.insertValue("A2", 2, 210, format_cell_id);
        //worksheet.saveDataToSheet(0);

        //worksheet.Add("Sheet2");

        //CellStyle cellStyle2 = new CellStyle()
        //{
        //    CellColor = "18FF6D",
        //    FontName = "Arial",
        //    FontSize = 10
        //};
        //uint format_cell_id2 = worksheet.insertStyle(cellStyle2);

        //worksheet.insertValue("B3", 3, " some more data ", format_cell_id2);
        //worksheet.insertValue("B4", 4, 30, format_cell_id2);
        //worksheet.saveDataToSheet(1);

        //workbook.Save(filePath);

        //*****************************************************************

        //for creating an Excel file with worksheets containg some data and images 
        //Workbook workbook = new Workbook();
        //Worksheet worksheet = new Worksheet(workbook);

        //CellStyle cellStyle = new CellStyle()
        //{
        //    CellColor = "18FF6D",
        //    FontName = "Calibri",
        //    FontSize = 22
        //};
        //uint format_cell_id = worksheet.insertStyle(cellStyle);

        //worksheet.insertValue("A1", 1, " some data ", format_cell_id);
        //worksheet.insertValue("A2", 2, 210, format_cell_id);
        //worksheet.saveDataToSheet(0);

        //Image img = new Image(workbook);
        //img.Add(0, imagePath, 6, 1, 8, 3);

        //worksheet.Add("Sheet2");

        //CellStyle cellStyle2 = new CellStyle()
        //{
        //    CellColor = "18FF6D",
        //    FontName = "Arial",
        //    FontSize = 10
        //};
        //uint format_cell_id2 = worksheet.insertStyle(cellStyle2);

        //worksheet.insertValue("B3", 3, " some more data ", format_cell_id2);
        //worksheet.insertValue("B4", 4, 30, format_cell_id2);
        //worksheet.saveDataToSheet(1);

        //Image img2 = new Image(workbook);
        //img2.Add(1, imagePath2,8,2,12,5);

        //workbook.Save(filePath);

        //Console.WriteLine(workbook.GetCellValue("Sheet1","A1"));
        //Image image = new Image(workbook);
        //Console.WriteLine(image.ExtractImagesFromWorkSheet(0));
        //Console.WriteLine(image.GetImagesCountFromWorkBook);
        //Worksheet worksheet = new Worksheet(workbook);

        //CellStyle cellStyle = new CellStyle()
        //{
        //    CellColor = "18FF6D",
        //    FontName = "Calibri",
        //    FontSize = 22
        //}; 

        //uint format_cell_id = worksheet.insertStyle(cellStyle);

        //worksheet.insertValue("A1", 1, " some data ", format_cell_id);
        //worksheet.insertValue("A2", 2, 210, format_cell_id);
        //worksheet.saveDataToSheet(0);

        //Image img = new Image(workbook);
        //img.Add(0, imagePath, 6, 1, 8, 3);

        //workbook.Save(filePath);

        //*****************************************************************

        // for image extraction as a tream
        //Workbook workbook = new Workbook(filePath);
        //Image image = new Image(workbook);
        //List<Stream> imageStream = image.ExtractImagesFromWorkSheet(0);
        //Console.WriteLine(imageStream);

        //*****************************************************************

        //for getting image count from a all worksheets in a workbook
        //Workbook workbook = new Workbook(filePath);
        //Image image = new Image(workbook);
        //int ImagesCount = image.GetImagesCountFromWorkBook;
        //Console.WriteLine(ImagesCount);

        //*****************************************************************

        //for getting cell's value
        //Workbook workbook = new Workbook(filePath);
        //Console.WriteLine(workbook.GetCellValue("Sheet1", "A1"));


        //*****************************************************************

        //for deleting value from cell
        //Workbook workbook = new Workbook(filePath);
        //workbook.DeleteTextFromCell("Sheet1", "A", 1);
        //workbook.Save(filePath);

        //*****************************************************************

        //for deleting a worksheet from a workbook
        Workbook workbook = new Workbook(filePath);
        workbook.DeleteWorksheet("Sheet2");
        workbook.Save(filePath);
    }
}