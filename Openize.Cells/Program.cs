

using Openize.Cells;
using System;

class Program
{
    static void Main()
    {

        string outputDirectory = "Z:\\Downloads\\";

        string filePath = "Z:\\Downloads\\testFile1.xlsx";

        var existingSheetName = "TestSheet";
        var newSheetName = "RenameSheetName";

        


        using (Workbook wb = new Workbook(filePath))
        {
            wb.SetSheetVisibility("RenameSheetName", SheetVisibility.Visible);
            wb.Save(filePath);
        }


        //using (Workbook wb = new Workbook(filePath))
        //{
        //    wb.ReorderSheets("RenameSheetName", 0);
        //    wb.Save(filePath);
        //}

        //using (Workbook wb = new Workbook(filePath))
        //{
        //    wb.CopySheet("RenameSheetName", "FahadCopySheet1");
        //    wb.Save(filePath);
        //}

        //using (Workbook wb = new Workbook(filePath))
        //{
        //    wb.RenameSheet(existingSheetName, newSheetName);
        //    wb.Save(filePath);
        //}

        //using (Workbook wb = new Workbook(filePath))
        //{
        //    Worksheet firstSheet = wb.Worksheets[0];

        //    string cellReference = "B10";
        //    Cell cell = firstSheet.Cells[cellReference];

        //    cell.PutValue("randomValue111");

        //    Comment myComment = new Comment("John Doe", "This is a test comment.");

        //    firstSheet.AddComment(cellReference, myComment);

        //    wb.Save(filePath);
        //}

        //using (Workbook wb = new Workbook(filePath))
        //{
        //    Worksheet firstSheet = wb.Worksheets[0];
        //    var range = firstSheet.GetRange("A1", "A10");

        //    firstSheet.CopyRange(range, "D1");




        //    wb.Save(filePath);

        //}



        //using (Workbook wb = new Workbook(filePath))
        //{


        //    Worksheet firstSheet = wb.Worksheets[0];

        //    uint startRowIndex = 5;
        //    uint numberOfRows = 3;

        //    firstSheet.InsertRows(startRowIndex, numberOfRows);

        //    int rowsCount = firstSheet.GetRowCount();

        //    Console.WriteLine("Rows Count=" + rowsCount);

        //    wb.Save(filePath);
        //}

        //using (Workbook wb = new Workbook(filePath))
        //{
        //    Worksheet firstSheet = wb.Worksheets[0];

        //    var range = firstSheet.GetRange("A1", "B10");
        //    Console.WriteLine($"column count: {range.ColumnCount}");
        //    Console.WriteLine($"row count: {range.RowCount}");

        //    range.SetValue("Hello");

        //    wb.Save(filePath);
        //}

        //using (Workbook wb = new Workbook(filePath))
        //{
        //    string startColumn = "B";
        //    int numberOfColumns = 3;


        //    Worksheet firstSheet = wb.Worksheets[0];


        //    int columnsCount = firstSheet.GetColumnCount();
        //    firstSheet.InsertColumns(startColumn, numberOfColumns);

        //    Console.WriteLine("Columns Count=" + columnsCount);

        //    wb.Save(filePath);
        //}

        //using (Workbook wb = new Workbook(filePath))
        //{


        //    Worksheet firstSheet = wb.Worksheets[0];

        //    uint startRowIndex = 5; 
        //    uint numberOfRows = 3; 

        //    firstSheet.InsertRows(startRowIndex,numberOfRows);

        //    int rowsCount = firstSheet.GetRowCount();

        //    Console.WriteLine("Rows Count=" + rowsCount);

        //    wb.Save(filePath);
        //}

        //using (Workbook wb = new Workbook(filePath))
        //{


        //    Worksheet firstSheet = wb.Worksheets[0];

        //    uint startRowIndex = 5; // Example: Start inserting rows from position 5
        //    uint numberOfRows = 3;  // Example: Insert 3 new rows
        //                            //firstSheet.InsertRows(startRowIndex,numberOfRows);

        //    // Assuming 'worksheet' is an instance of your worksheet class
        //    string startColumn = "I";  // Example: Insert a column at column "B"
        //                               //firstSheet.InsertColumn(startColumn);  // Inserts one column
        //                               // or
        //                               //firstSheet.InsertColumns(startColumn, 2);  // Inserts two columns starting from column "B"
        //                               //firstSheet.HideColumn("B");
        //    firstSheet.UnhideColumn("B");
        //    //firstSheet.UnhideRows(6,3);

        //    wb.Save(filePath);
        //}


        //using (Workbook wb = new Workbook())
        //{

        //    Worksheet firstSheet = wb.Worksheets[0];
        //    uint rowIndex = 5; // Example: Hide row 5
        //    firstSheet.HideRow(rowIndex);
        //    wb.Save(filePath);
        //}
        //using (Workbook wb = new Workbook())
        //{

        //    Worksheet firstSheet = wb.Worksheets[0];   
        //    string columnName = "B"; // Example: Hide column B
        //    firstSheet.HideColumn(columnName);
        //    wb.Save(filePath);
        //}



        //using (Workbook wb = new Workbook(filePath))
        //{
        //    Worksheet firstSheet = wb.Worksheets[0];
        //    var cellReference = "A1";
        //    // Retrieve the validation rule for a specific cell
        //    ValidationRule rule = firstSheet.GetValidationRule(cellReference);

        //    if (rule != null)
        //    {
        //        // Display information about the retrieved validation rule
        //        Console.WriteLine($"Validation Type: {rule.Type}");
        //        Console.WriteLine($"Error Message: {rule.ErrorMessage}");
        //    }
        //    else
        //    {
        //        Console.WriteLine("No validation rule found for the specified cell.");
        //    }
        //}



        //using (workbook wb = new workbook()) // creating a new workbook
        //{
        //    //console.writeline("styleindex = " + styleindex);
        //    worksheet firstsheet = wb.worksheets[0];

        //    //var range = firstsheet.getrange("a1", "b10");
        //    //console.writeline($"column count: {range.columncount}");
        //    //console.writeline($"row count: {range.rowcount}");

        //    //range.setvalue("hello2");
        //    //range.clearcells();
        //    //range.mergecells();

        //    //validationrule listrule = new validationrule(new[] { "option1", "option2", "option3" });

        //    validationrule listrule = new validationrule(new[] { "apple", "banana", "orange" })
        //    {
        //        errortitle = "invalid entry",
        //        errormessage = "please choose a fruit from the list: apple, banana, or orange."
        //    };


        //    validationrule numberrule = new validationrule(validationtype.wholenumber, 1, 100);
        //    // adding dropdown list validation to the range

        //    validationrule customformulavalidation = new validationrule("=and(a1>0, a1<100)");

        //    firstsheet.applyvalidation("b1", customformulavalidation); // applying to cell b1

        //    firstsheet.applyvalidation("a1", listrule);

        //    firstsheet.applyvalidation("a2", numberrule);
        //    //firstsheet.adddropdownlistvalidation("a1", dropdownoptions);

        //    // save the workbook
        //    wb.save(filepath);
        //}


        //using (Workbook wb = new Workbook(filePath))
        //{
        //    wb.UpdateDefaultStyle("Arial", 12, "A02000");
        //    //Console.WriteLine("Style ID = " + styleId);
        //    wb.Save();
        //}



        //using (Workbook wb = new Workbook(filePath)) // Open Existing workbook
        //{

        //    foreach (var worksheet in wb.Worksheets)
        //    {
        //        if (worksheet.IsProtected())
        //        {
        //            Console.WriteLine("Protect Sheet Name = " + worksheet.Name);
        //            worksheet.UnprotectSheet();
        //        }
        //    }
        //    // Save the workbook
        //    wb.Save(filePath);
        //}

        //using (Workbook wb = new Workbook(filePath)) // Open Existing workbook
        //{
        //    int i = 0;
        //    foreach (var worksheet in wb.Worksheets)
        //    {
        //        i++;
        //        worksheet.Name = $"Fahad{i}";
        //        Console.WriteLine(worksheet.Name);
        //    }
        //    // Save the workbook
        //    wb.Save(filePath);
        //}



        //using (Workbook wb = new Workbook()) // Creating a new workbook
        //{

        //    Worksheet firstSheet = wb.Worksheets[0];
        //    firstSheet.SetRowHeight(1, 40);      // Set height of row 1 to 40 points
        //    firstSheet.SetColumnWidth("B", 75);
        //    // Put values into cells
        //    Cell cellA1 = firstSheet.Cells["A1"];
        //    cellA1.PutValue("aaa A1");

        //    // Repeat the process for other cells as needed
        //    Cell cellB2 = firstSheet.Cells["B2"];
        //    cellB2.PutValue("Styled Text");
        //    // Save the workbook
        //    wb.Save(filePath);
        //}


        //using (Workbook wb = new Workbook()) // Creating a new workbook
        //{
        //Console.WriteLine("styleIndex = " + styleIndex);
        //    Worksheet firstSheet = wb.Worksheets[0];
        // Put values into cells
        //    Cell cellA1 = firstSheet.Cells["A1"];
        //    cellA1.PutValue("aaa A1");
        //    firstSheet.ProtectSheet("a2387ass");
        // Save the workbook
        //    wb.Save(filePath);
        //}

        //Workbook wb = new Workbook(filePath);

        //var worksheet = wb.Worksheets[0];
        //var images = worksheet.ExtractImages();
        //if (!Directory.Exists(outputDirectory))
        //{
        //    Directory.CreateDirectory(outputDirectory);
        //}

        //foreach (var image in images)
        //{
        //    // Construct file path
        //    var outputFilePath = Path.Combine(outputDirectory, $"Image_{Guid.NewGuid()}.{image.Extension}");

        //    // Save image data to file
        //    using (var fileStream = File.Create(outputFilePath))
        //    {
        //        image.Data.CopyTo(fileStream);
        //    }
        //}




        //using (var workbook = new Workbook())
        //{
        //    var firstSheet = workbook.Worksheets[0];
        //    var image = new Image("/Users/fahadadeelqazi/Downloads/ImageCells.png");
        //    Console.WriteLine("Image = " + image.Path);
        //    firstSheet.AddImage(image, 6, 1, 8, 3);

        //    var secondSheet = workbook.AddSheet("fahad");

        //    var image1 = new Image("/Users/fahadadeelqazi/Downloads/ImageCells.png");
        //    Console.WriteLine("Image = " + image.Path);
        //    secondSheet.AddImage(image1, 6, 1, 8, 3);

        //    workbook.Save(filePath);
        //}



        //using (Workbook wb = new Workbook()) // Creating a new workbook
        //{
        //    // Create a style with Calibri font, size 11, and red color
        //    uint styleIndex = wb.CreateStyle("Arial", 11, "FF0000");
        //    uint styleIndex2 = wb.CreateStyle("Times New Roman", 12, "000000");

        //    //Console.WriteLine("styleIndex = " + styleIndex);
        //    Worksheet firstSheet = wb.Worksheets[0];

        //    // Put values into cells
        //    Cell cellA1 = firstSheet.Cells["A1"];
        //    cellA1.PutValue("aaa A1");

        //    // Apply the style to the cell
        //    cellA1.ApplyStyle(styleIndex);

        //    // Repeat the process for other cells as needed
        //    Cell cellB2 = firstSheet.Cells["B2"];
        //    cellB2.PutValue("Styled Text");
        //    cellB2.ApplyStyle(styleIndex2);

        //    // Save the workbook
        //    wb.Save(filePath);
        //}

        // Example code for Merge Cells

        //using (var workbook = new Workbook())
        //{
        //    var firstSheet = workbook.Worksheets[0];
        //    firstSheet.MergeCells("A1", "C1"); // Merge cells from A1 to C1

        //    // Add value to the top-left cell of the merged area
        //    var topLeftCell = firstSheet.Cells["A1"];
        //    topLeftCell.PutValue("This is a merged cell");

        //    workbook.Save(filePath);
        //}

        // Example for setting Default Style for the whole workbook and some custom styles for cells.
        //using (var workbook = new Workbook())
        //{
        //    // Update default style and create new styles
        //    workbook.UpdateDefaultStyle("Times New Roman", 11, "000000");
        //    uint headerStyleIndex = workbook.CreateStyle("Arial", 15, "000000"); // Black for headers
        //    uint evenRowStyleIndex = workbook.CreateStyle("Arial", 12, "FF0000"); // Red for even
        //    uint oddRowStyleIndex = workbook.CreateStyle("Calibri", 12, "0000FF"); // Blue for odd

        //    var firstSheet = workbook.Worksheets[0];

        //    // Header row
        //    string[] headers = { "Student ID", "Student Name", "Course", "Grade" };
        //    for (int col = 0; col < headers.Length; col++)
        //    {
        //        string cellAddress = $"{(char)(65 + col)}1"; // A1, B1, etc.
        //        Cell cell = firstSheet.Cells[cellAddress];
        //        cell.PutValue(headers[col]);
        //        cell.ApplyStyle(headerStyleIndex);
        //    }

        //    // Data rows
        //    int rowCount = 10;
        //    for (int row = 2; row <= rowCount + 1; row++) // Starting from row 2 because row 1 is header
        //    {
        //        for (int col = 0; col < headers.Length; col++)
        //        {
        //            string cellAddress = $"{(char)(65 + col)}{row}"; // Converts 0 to A, 1 to B, etc., and appends row number
        //            Cell cell = firstSheet.Cells[cellAddress];

        //            // Sample data generation logic. 
        //            switch (col)
        //            {
        //                case 0: // Student ID
        //                    cell.PutValue($"ID{row - 1}");
        //                    break;
        //                case 1: // Student Name
        //                    cell.PutValue($"Student {row - 1}");
        //                    break;
        //                case 2: // Course
        //                    cell.PutValue($"Course {(row - 1) % 5 + 1}");
        //                    break;
        //                case 3: // Grade
        //                    cell.PutValue($"Grade {((row - 1) % 3) + 'A'}");
        //                    break;
        //            }

        //            // Apply different styles for even and odd rows
        //            cell.ApplyStyle((row % 2 == 0) ? evenRowStyleIndex : oddRowStyleIndex);
        //        }
        //    }

        //    workbook.Save(filePath);
        //}

        // Properties example
        //using (var workbook = new Workbook())
        //{
        //    Worksheet firstSheet = workbook.Worksheets[0];
        //    Cell cellA1 = firstSheet.Cells["A1"];
        //    cellA1.PutValue("Text A1");

        //    Cell cellA2 = firstSheet.Cells["A2"];
        //    cellA2.PutValue("Text A2");
        //    // Set new properties
        //    var newProperties = new BuiltInDocumentProperties
        //    {
        //        Author = "Fahad Adeel",
        //        Title = "Sample Workboo1",
        //        CreatedDate = DateTime.Now,
        //        ModifiedBy = "Fahad",
        //        ModifiedDate = DateTime.Now.AddHours(1),
        //        Subject = "Testing Subject"
        //    };

        //    workbook.BuiltinDocumentProperties = newProperties;

        //    workbook.Save(filePath);
        //}

        //static void DisplayProperties(BuiltInDocumentProperties properties)
        //{
        //    Console.WriteLine($"Author: {properties.Author}");
        //    Console.WriteLine($"Title: {properties.Title}");
        //    Console.WriteLine($"Created Date: {properties.CreatedDate}");
        //    Console.WriteLine($"Modified By: {properties.ModifiedBy}");
        //    Console.WriteLine($"Modified Date: {properties.ModifiedDate}");
        //    Console.WriteLine("=================================");
        //}

        //Create a new workbook
        //Scenario 1: Create a new workbook and save it to a specific file path.
        //using (Workbook wb = new Workbook())
        //{
        //    Worksheet firstSheet = wb.Worksheets[0];

        //    // Put values into cells
        //    Cell cellA1 = firstSheet.Cells["A1"];
        //    cellA1.PutValue("aaa A1");

        //    var newSheet = wb.AddSheet("FahadSheet");
        //    Cell cellB1 = newSheet.Cells["B1"];
        //    cellB1.PutValue("bbb B1");
        //    wb.Save(filePath);
        //}

        //Example code for adding formula to cell.
        //using (Workbook wb = new Workbook())
        //{
        //    Worksheet firstSheet = wb.Worksheets[0];

        //    Random rand = new Random();
        //    for (int i = 1; i <= 10; i++)
        //    {
        //        string cellReference = $"A{i}";
        //        Cell cell = firstSheet.Cells[cellReference];
        //        double randomValue = rand.Next(1, 100); // Generate random number between 1 and 100
        //        cell.PutValue(randomValue); // Put random number into cell
        //    }

        //    Cell cellA11 = firstSheet.Cells["A11"];
        //    cellA11.PutFormula("SUM(A1:A10)"); // Putting a formula into cell A11 to sum A1 to A10
        //    wb.Save(filePath); // Saving the workbook

        //    Console.WriteLine("VAAA=" +cellA11.GetValue());

        //}

        //using (Workbook wb = new Workbook(filePath))
        //{
        //    Worksheet firstSheet = wb.Worksheets[0];


        //    Cell cellA11 = firstSheet.Cells["A11"];


        //    Console.WriteLine("VAAA11=" + cellA11.GetValue());

        //}



        // Output the value stored in a cell
        //using (Workbook wb = new Workbook(filePath))
        //{
        //    Worksheet firstSheet = wb.Worksheets[0];
        //    Cell cellA1 = firstSheet.Cells["A10"];
        //    Console.WriteLine(cellA1.GetDataType());
        //    string value = cellA1.GetValue();

        //    Console.WriteLine(value); // Output the value stored in cell A1
        //}

        // Remove Worksheet by name

        //using (Workbook wb = new Workbook(filePath))
        //{
        //    bool isRemoved = wb.RemoveSheet("FahadSheet");
        //    if (isRemoved)
        //    {
        //        // Save the workbook if the sheet is successfully removed
        //        wb.Save();
        //    }
        //}

        //Scenario 2: Open an existing workbook, modify it, and save changes back to the original file.
        //using (Workbook wb = new Workbook(filePath))
        //{

        //    Console.WriteLine(wb.Worksheets.Count);
        //    Worksheet firstSheet = wb.Worksheets[1];
        //    Cell cell = firstSheet.Cells["D2"];
        //    cell.PutValue("Fahad");
        //    wb.Save();
        //}

        // Scenario 3: Open an existing workbook, modify it, and save to a MemoryStream.
        //using (Workbook wb = new Workbook(filePath))
        //{
        //    var newSheet = wb.AddNewSheet("NewSheetName2");
        //    Cell cell = newSheet["A1"];
        //    cell.PutValue("Hello from another new sheet!");
        //    using (MemoryStream ms = new MemoryStream())
        //    {
        //        wb.Save(ms);

        //        // Do something with the MemoryStream, such as sending it to a client, etc.
        //    }
        //}
    }
}


