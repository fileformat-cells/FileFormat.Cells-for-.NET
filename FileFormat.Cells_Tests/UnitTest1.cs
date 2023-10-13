namespace FileFormat.Cells_Tests;
using FileFormat.Cells;

[TestClass]
public class UnitTest1
{

    string testFilePath = "/Users/fahadadeelqazi/Downloads/test_fahad_new_protected111.xlsx";
    private string imageInputPath = "/Users/fahadadeelqazi/Downloads/ImageCells.png";
    string outputDirectory = "/Users/fahadadeelqazi/Downloads/";

    [TestInitialize]
    public void Setup()
    {

        if (File.Exists(testFilePath))
            File.Delete(testFilePath);

        //if (Directory.Exists(outputDirectory))
        //    Directory.Delete(outputDirectory, true);

        // Initialize a test workbook with 3 worksheets
        using (Workbook wb = new Workbook())
        {
            wb.AddSheet("Sheet2");
            wb.AddSheet("Sheet3");
            wb.Save(testFilePath);
        }
    }

    [TestMethod]
    public void Test_AddSheet()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            var newSheet = wb.AddSheet("TestSheet");
            Assert.AreEqual("TestSheet", newSheet.Name);
        }
    }

    [TestMethod]
    public void Test_RemoveSheet()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            var result = wb.RemoveSheet("Sheet2");
            Assert.IsTrue(result);
            Assert.IsFalse(wb.Worksheets.Any(ws => ws.Name == "Sheet2"));
        }
    }

    [TestMethod]
    public void Test_Save()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            wb.AddSheet("SaveTestSheet");
            wb.Save();

            using (Workbook testWb = new Workbook(testFilePath))
            {
                Assert.IsTrue(testWb.Worksheets.Any(ws => ws.Name == "SaveTestSheet"));
            }
        }
    }

    [TestMethod]
    public void Test_SaveAs()
    {
        string newFilePath = "TestWorkbook_SaveAs.xlsx";

        using (Workbook wb = new Workbook(testFilePath))
        {
            wb.AddSheet("SaveAsTestSheet");
            wb.Save(newFilePath);
        }

        Assert.IsTrue(File.Exists(newFilePath));

        using (Workbook testWb = new Workbook(newFilePath))
        {
            Assert.IsTrue(testWb.Worksheets.Any(ws => ws.Name == "SaveAsTestSheet"));
        }

        if (File.Exists(newFilePath))
            File.Delete(newFilePath);
    }

    [TestMethod]
    public void Test_UpdateDefaultStyle()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            wb.UpdateDefaultStyle("Arial", 12, "#FF0000");
        }
    }

    [TestMethod]
    public void Test_CreateStyle()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            var styleId = wb.CreateStyle("Arial", 12, "FF0000");
            Console.Write(styleId);
            Assert.IsTrue(styleId > 0);
        }
    }

    [TestMethod]
    public void Test_BuiltinDocumentProperties()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            var props = wb.BuiltinDocumentProperties;
            props.Author = "TestAuthor";
            props.Title = "TestTitle";
            wb.BuiltinDocumentProperties = props;
            wb.Save();
        }

        using (Workbook wb = new Workbook(testFilePath))
        {
            var props = wb.BuiltinDocumentProperties;
            Assert.AreEqual("TestAuthor", props.Author);
            Assert.AreEqual("TestTitle", props.Title);
        }
    }

    [TestMethod]
    public void Test_ProtectSheet()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            var testWorksheet = wb.Worksheets.First();
            testWorksheet.UnprotectSheet();   // Ensure the sheet is unprotected before testing
            Assert.IsFalse(testWorksheet.IsProtected(), "Sheet should not be protected before testing ProtectSheet.");

            // Act
            testWorksheet.ProtectSheet("123");

            // Assert
            Assert.IsTrue(testWorksheet.IsProtected(), "Sheet should be protected after calling ProtectSheet.");
        }
    }

    [TestMethod]
    public void Test_UnprotectSheet()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            var testWorksheet = wb.Worksheets.First();
            testWorksheet.ProtectSheet("123");   // Ensure the sheet is protected before testing
            Assert.IsTrue(testWorksheet.IsProtected(), "Sheet should be protected before testing UnprotectSheet.");

            // Act
            testWorksheet.UnprotectSheet();

            // Assert
            Assert.IsFalse(testWorksheet.IsProtected(), "Sheet should not be protected after calling UnprotectSheet.");
        }
    }

    [TestMethod]
    public void Test_SetCellValue()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            Worksheet firstSheet = wb.Worksheets[0];
            Cell cellA10 = firstSheet.Cells["A10"];

            // Act
            cellA10.PutValue("TestValue");
            wb.Save();

            // Assert
            string savedValue;
            using (Workbook testWb = new Workbook(testFilePath))
            {
                savedValue = testWb.Worksheets[0].Cells["A10"].GetValue();
            }
            Assert.AreEqual("TestValue", savedValue);
        }
    }

    [TestMethod]
    public void Test_GetCellValue()
    {
        string expectedValue = "ExpectedTestValue";

        using (Workbook wb = new Workbook(testFilePath))
        {
            Worksheet firstSheet = wb.Worksheets[0];
            Cell cellA10 = firstSheet.Cells["A10"];
            cellA10.PutValue(expectedValue);
            wb.Save();
        }

        using (Workbook wb = new Workbook(testFilePath))
        {
            Worksheet firstSheet = wb.Worksheets[0];
            Cell cellA10 = firstSheet.Cells["A10"];

            // Act
            string retrievedValue = cellA10.GetValue();

            // Assert
            Assert.AreEqual(expectedValue, retrievedValue);
        }
    }

    [TestMethod]
    public void Test_PutFormula()
    {
        string expectedFormula = "SUM(A1:A10)";

        using (Workbook wb = new Workbook(testFilePath))
        {
            Worksheet firstSheet = wb.Worksheets[0];

            // Fill cells A1 to A10 with data
            for (int i = 1; i <= 10; i++)
            {
                string cellReference = $"A{i}";
                Cell cell = firstSheet.Cells[cellReference];
                double value = i; // or any other data you'd like to put
                cell.PutValue(value);
            }

            Cell cellA11 = firstSheet.Cells["A11"];
            cellA11.PutFormula(expectedFormula);
            wb.Save();
        }

        using (Workbook wb = new Workbook(testFilePath))
        {
            Worksheet firstSheet = wb.Worksheets[0];
            Cell cellA11 = firstSheet.Cells["A11"];
            string actualFormula = cellA11.GetFormula();

            Assert.AreEqual(expectedFormula, actualFormula);

        }
    }

    [TestMethod]
    public void Test_MergeCells()
    {
        string expectedValue = "This is a merged cell";

        using (Workbook wb = new Workbook(testFilePath))
        {
            Worksheet firstSheet = wb.Worksheets[0];

            // Merge cells from A1 to C1
            firstSheet.MergeCells("A1", "C1");

            // Add value to the top-left cell of the merged area
            Cell topLeftCell = firstSheet.Cells["A1"];
            topLeftCell.PutValue(expectedValue);

            wb.Save();
        }

        using (Workbook wb = new Workbook(testFilePath))
        {
            Worksheet firstSheet = wb.Worksheets[0];
            Cell mergedTopLeftCell = firstSheet.Cells["A1"];

            // Check the value of the top-left cell of the merged area
            Assert.AreEqual(expectedValue, mergedTopLeftCell.GetValue());

        }
    }

    [TestMethod]
    public void Test_AddImages()
    {
        using (Workbook wb = new Workbook())
        {
            var firstSheet = wb.Worksheets[0];
            var image = new Image(imageInputPath);
            firstSheet.AddImage(image, 6, 1, 8, 3);

            var secondSheet = wb.AddSheet("fahad");
            var image1 = new Image(imageInputPath);
            secondSheet.AddImage(image1, 6, 1, 8, 3);

            wb.Save(testFilePath);
        }
    }

    [TestMethod]
    public void Test_ExtractImages()
    {
        using (Workbook wb = new Workbook(testFilePath))
        {
            var worksheet = wb.Worksheets[0];
            var images = worksheet.ExtractImages();

            Assert.IsTrue(images.Any(), "No images extracted from the worksheet.");

            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            foreach (var image in images)
            {
                var outputFilePath = Path.Combine(outputDirectory, $"Image_{Guid.NewGuid()}.{image.Extension}");

                using (var fileStream = File.Create(outputFilePath))
                {
                    image.Data.CopyTo(fileStream);
                }

                // Assert that the image has been saved correctly
                Assert.IsTrue(File.Exists(outputFilePath), $"Image not found at {outputFilePath}");
            }
        }
    }








    [TestCleanup]
    public void Cleanup()
    {
        // Clean up test artifacts
        //if (File.Exists(testFilePath))
            //File.Delete(testFilePath);

        
    }

}
