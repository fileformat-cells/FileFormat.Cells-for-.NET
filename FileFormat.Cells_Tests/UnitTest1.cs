namespace FileFormat.Cells_Tests;
using FileFormat.Cells;
using FileFormat.Cells.Image;
using FileFormat.Cells.Properties;

[TestClass]
public class UnitTest1
{

    private static string testDir = "/Users/fahadadeelqazi/Projects/Aspose/FileFormat.Cells/TestSpreadSheets/";
    private static string processedDir = "/Users/fahadadeelqazi/Projects/Aspose/FileFormat.Cells/TestSpreadSheets/ProcessedSpreadSheets/";
    private static string testSpreadsheet = "UbuntuSoftwareCenter";


    [TestMethod]
    public void TestCreateNSave()
    {
        try
        {
            // creates a workbook with one default worksheet.
            Workbook workbook = new Workbook();
            workbook.Save(processedDir + "Created_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".xlsx");

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while saving the workbook: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }

    [TestMethod]
    public void TestWorkSheetCreateNSave()
    {
        try
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet(workbook);
            worksheet.Add("sheet2");
            workbook.Save(processedDir + "Created_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".xlsx");

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while saving the workbook: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
    [TestMethod]
    public void TestInsertTextIntoWorksheet()
    {
        try
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet(workbook);

            worksheet.insertValue("A10", 10, "some data", 0);
            worksheet.saveDataToSheet(0);
            workbook.Save(processedDir + "Created_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".xlsx");

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while saving the workbook: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
    [TestMethod]
    public void TestInsertTextWithStyleIntoWorksheet()
    {
        try
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet(workbook);

            CellStyle cellStyle = new CellStyle()
            {
                CellColor = "18FF6D",
                FontName = "Calibri",
                FontSize = 22
            };

            uint format_cell_id = worksheet.insertStyle(cellStyle);

            worksheet.insertValue("A10", 10, "some data", format_cell_id);
            worksheet.saveDataToSheet(0);
            workbook.Save(processedDir + "Created_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".xlsx");

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while saving the workbook: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }

    [TestMethod]
    public void TestApplyFonntStyleIntoWorkbook()
    {
        try
        {
            Workbook workbook = new Workbook();
            workbook.ApplyFontStyle("Arial", 14);
            Worksheet worksheet = new Worksheet(workbook);

            worksheet.insertValue("A10", 10, "some data", 0);
            worksheet.saveDataToSheet(0);
            workbook.Save(processedDir + "Created_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".xlsx");

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while saving the workbook: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }

    
    [TestMethod]
    public void TestDeleteWorksheet()
    {
        try
        {
            string path = testDir + "test.xlsx";
            Workbook workbook = new Workbook(path);
            // Delete worksheet named custom2
            workbook.DeleteWorksheet("custom2");
            workbook.Save(processedDir + "Created_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".xlsx");

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while saving the workbook: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
    [TestMethod]
    public void TestAddImage()
    {
        try
        {
            Workbook workbook = new Workbook(); 
            Image img = new Image(workbook);
            img.Add(0, testDir+"pic.png", 6, 1, 8, 3);
            workbook.Save(processedDir + "Created_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".xlsx");

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while adding the images into the workbook: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
    [TestMethod]
    public void TestExtractImagesFromWorkSheet()
    {
        try
        {
            string path = testDir + "test.xlsx";
            Workbook workbook = new Workbook(path);
            Image image = new Image(workbook);
            image.ExtractImagesFromWorkSheet(0);

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while extracting the images from the workbook: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
    [TestMethod]
    public void TestGetImagesCountFromWorkBook()
    {
        try
        {
            string path = testDir + "test.xlsx";
            Workbook workbook = new Workbook(path);
            Image image = new Image(workbook);
            int ImagesCount = image.GetImagesCountFromWorkBook;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while getting the total number of images in a Workbook: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
    [TestMethod]
    public void TestGetCellValue()
    {
        try
        {
            string path = testDir + "test.xlsx";
            Workbook workbook = new Workbook(path);
            workbook.GetCellValue("Sheet1", "A1");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while reading the cell's value from the worksheet: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
    [TestMethod]
    public void TestDeleteTextFromCell()
    {
        try
        {
            string path = testDir + "test.xlsx";
            Workbook workbook = new Workbook(path);
            workbook.DeleteTextFromCell("Sheet1", "A",1);
            workbook.Save(processedDir + "Created_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while reading the cell's value from the worksheet: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
    [TestMethod]
    public void TestGetBuiltInDocumentProperties()
    {
        try {
            string path = testDir + "test.xlsx";

            BuiltInDocumentProperties prop;
            using (Workbook workbook = new Workbook(path))
            {
                workbook.Save(path);
                prop = workbook.BuiltinDocumentProperties;
            }
            string author = prop.Author;
            DateTime creationDate = prop.CreatedDate;
            string modifier = prop.ModifiedBy;
            DateTime modificationDate = prop.ModifiedDate;
            string title = prop.Title;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred in built-in props: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
        }

    }

}
