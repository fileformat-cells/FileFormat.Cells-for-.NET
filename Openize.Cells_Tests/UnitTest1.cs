
using Openize.Cells;
namespace Openize.Cells_Tests
{
    [TestClass]
    public class UnitTest1
    {

        string testFilePath = "Z:\\Downloads\\test_fahad_new_protected_image1.xlsx";
        private string imageInputPath = "Z:\\Downloads\\ImageCells.png";
        string outputDirectory = "Z:\\Downloads\\";

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
                var firstSheet = wb.Worksheets.First();
                Cell cellA1 = firstSheet.Cells["A1"];
                Cell cellB1 = firstSheet.Cells["B1"];

                
                cellA1.PutValue("A1");
                cellB1.PutValue("B1");

                wb.AddSheet("Sheet2");
                wb.AddSheet("Sheet3");
                wb.Save(testFilePath);
            }
            
        }


        [TestMethod]
        public void Test_InsertRows()
        {
            Setup();  // Create a new workbook

            uint startRowIndex = 5;
            uint numberOfRows = 3;

            // Open the workbook and get the initial row count
            int initialRowCount;
            using (Workbook wb = new Workbook(testFilePath))
            {
                var worksheet = wb.Worksheets.First();
                initialRowCount = worksheet.GetRowCount();
            }

            // Re-open the workbook, insert rows, and save
            using (Workbook wb = new Workbook(testFilePath))
            {
                var worksheet = wb.Worksheets.First();

                // Act - Insert rows
                worksheet.InsertRows(startRowIndex, numberOfRows);
                wb.Save();
            }

            // Re-open the workbook and verify the row count
            using (Workbook wb = new Workbook(testFilePath))
            {
                var worksheet = wb.Worksheets.First();

                // Assert
                int newRowCount = worksheet.GetRowCount();
                Assert.AreEqual(initialRowCount + numberOfRows, newRowCount, "Incorrect number of rows after insertion.");
            }
        }

        [TestMethod]
        public void Test_InsertColumns()
        {
            Setup();  // Create a new workbook

            string startColumn = "B";
            int numberOfColumns = 3;

            // Open the workbook and get the initial column count
            int initialColumnCount;
            using (Workbook wb = new Workbook(testFilePath))
            {
                var worksheet = wb.Worksheets.First();
                initialColumnCount = worksheet.GetColumnCount();
                Console.WriteLine("initialColumnCount=" + worksheet.GetColumnCount());
            }

            // Re-open the workbook, insert columns, and save
            using (Workbook wb = new Workbook(testFilePath))
            {
                var worksheet = wb.Worksheets.First();

                // Act - Insert columns
                worksheet.InsertColumns(startColumn, numberOfColumns);
                
                wb.Save();
            }

            // Re-open the workbook and verify the column count
            using (Workbook wb = new Workbook(testFilePath))
            {
                var worksheet = wb.Worksheets.First();

                // Assert
                int newColumnCount = worksheet.GetColumnCount();
                Console.WriteLine("newColumnCount=" + newColumnCount);
                Assert.AreEqual(initialColumnCount, newColumnCount, "Incorrect number of columns after insertion.");
            }
        }

        [TestMethod]
        public void Test_HideRows()
        {
            Setup();  // Setup your test workbook

            uint startRowIndex = 5;
            uint numberOfRows = 3;

            using (Workbook wb = new Workbook(testFilePath))
            {
                var worksheet = wb.Worksheets.First();

                // Act - Hide the rows
                worksheet.HideRows(startRowIndex, numberOfRows);
                wb.Save();
            }

            // Assert
            using (Workbook wb = new Workbook(testFilePath))
            {
                var worksheet = wb.Worksheets.First();

                // Verify each row in the range is hidden
                for (uint rowIndex = startRowIndex; rowIndex < startRowIndex + numberOfRows; rowIndex++)
                {
                    Assert.IsTrue(worksheet.IsRowHidden(rowIndex), $"Row {rowIndex} should be hidden.");
                }
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
                var image = new Image(imageInputPath);

                worksheet.AddImage(image, 6, 1, 8, 3);

                wb.Save(testFilePath);

                Workbook wb1 = new Workbook(testFilePath);

                var worksheet1 = wb1.Worksheets[0];

                var images = worksheet1.ExtractImages();

                Assert.IsTrue(images.Any(), "No images extracted from the worksheet.");

                if (!Directory.Exists(outputDirectory))
                    Directory.CreateDirectory(outputDirectory);

                foreach (var extracted_image in images)
                {
                    var outputFilePath = Path.Combine(outputDirectory, $"Image_{Guid.NewGuid()}.{extracted_image.Extension}");

                    using (var fileStream = File.Create(outputFilePath))
                    {
                        extracted_image.Data.CopyTo(fileStream);
                    }

                    // Assert that the image has been saved correctly
                    Assert.IsTrue(File.Exists(outputFilePath), $"Image not found at {outputFilePath}");
                }
            }


        }

        [TestMethod]
        public void ListValidationRule_CreatesCorrectType()
        {
            var options = new[] { "Option1", "Option2", "Option3" };
            var rule = new ValidationRule(options);

            Assert.AreEqual(ValidationType.List, rule.Type);
            CollectionAssert.AreEqual(options, rule.Options);
        }

        [TestMethod]
        public void NumericValidationRule_CreatesCorrectTypeAndValues()
        {
            var minValue = 10.0;
            var maxValue = 100.0;
            var rule = new ValidationRule(ValidationType.Decimal, minValue, maxValue);

            Assert.AreEqual(ValidationType.Decimal, rule.Type);
            Assert.AreEqual(minValue, rule.MinValue);
            Assert.AreEqual(maxValue, rule.MaxValue);
        }

        [TestMethod]
        public void CustomFormulaValidationRule_CreatesCorrectFormula()
        {
            var formula = "=A1>0";
            var rule = new ValidationRule(formula);

            Assert.AreEqual(ValidationType.CustomFormula, rule.Type);
            Assert.AreEqual(formula, rule.CustomFormula);
        }

        [TestMethod]
        public void ApplyListValidation_And_Verify()
        {
            var expectedOptions = new[] { "Apple", "Banana", "Orange" };
            using (var workbook = new Workbook())
            {
                var worksheet = workbook.Worksheets[0];

                // Act: Apply a list validation rule to a cell
                var listRule = new ValidationRule(expectedOptions);
                worksheet.ApplyValidation("A1", listRule); // Applying to cell A1

                // Act: Save the workbook
                workbook.Save(testFilePath);
            }

            // Assert: Reopen the workbook and verify the validation rule
            using (var workbook = new Workbook(testFilePath))
            {
                var worksheet = workbook.Worksheets[0];
                var retrievedRule = worksheet.GetValidationRule("A1");

                Console.WriteLine("Expected Options: " + string.Join(", ", expectedOptions));
                Console.WriteLine("Retrieved Options: " + string.Join(", ", retrievedRule.Options));

                // Assert: Check that the retrieved rule matches what was applied
                Assert.IsNotNull(retrievedRule, "Validation rule should not be null.");
            
                Assert.AreEqual(ValidationType.List, retrievedRule.Type, "Validation type should be List.");

                // Verify each option individually
                var processedRetrievedOptions = retrievedRule.Options
                                            .Select(option => option.Trim(new char[] { ' ', '"' }))
                                            .ToArray();

                CollectionAssert.AreEqual(expectedOptions, processedRetrievedOptions, "Validation options do not match.");
            }
        }

    
        [TestCleanup]
        public void Cleanup()
        {
            //Clean up test artifacts
            if (File.Exists(testFilePath))
                File.Delete(testFilePath);

        }

    }
}
