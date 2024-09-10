using System;
using System.IO;
using CsvHelper;
using System.Globalization;
using System.Collections.Generic;
using Xceed.Words.NET;

class Program
{
    static void Main()
    {
        Console.WriteLine("Enter the path of the CSV file:");
        string csvFilePath = Console.ReadLine();

        Console.WriteLine("Enter the path where you want to save the Word document (with filename and .docx extension):");
        string docxFilePath = Console.ReadLine();

        if (!File.Exists(csvFilePath))
        {
            Console.WriteLine("CSV file not found. Please check the path and try again.");
            return;
        }

        try
        {
            ConvertCsvToDocx(csvFilePath, docxFilePath);
            Console.WriteLine($"CSV file successfully converted to Word document and saved at {docxFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static void ConvertCsvToDocx(string csvFilePath, string docxFilePath)
    {
        // Open the Word document
        using (DocX document = DocX.Create(docxFilePath))
        {
            // Add a title
            document.InsertParagraph("CSV Data").FontSize(20).Bold().SpacingAfter(20);

            // Read the CSV file
            using (var reader = new StreamReader(csvFilePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = new List<dynamic>();
                var headerRow = csv.Read();
                csv.ReadHeader();

                // Create a table with the header and data rows
                var table = document.AddTable(csv.Context.HeaderRecord.Length, 1);
                table.Alignment = Xceed.Document.NET.Alignment.center;
                table.Design = TableDesign.TableGrid;

                // Set the header row
                for (int i = 0; i < csv.Context.HeaderRecord.Length; i++)
                {
                    table.Rows[0].Cells[i].Paragraphs[0].Append(csv.Context.HeaderRecord[i]).Bold();
                }

                // Read the data rows
                while (csv.Read())
                {
                    var rowData = csv.GetRecord<dynamic>();
                    var dataRow = new List<string>();

                    foreach (var field in rowData)
                    {
                        dataRow.Add(field.Value);
                    }

                    // Insert each row into the table
                    var row = table.InsertRow();
                    for (int i = 0; i < dataRow.Count; i++)
                    {
                        row.Cells[i].Paragraphs[0].Append(dataRow[i]);
                    }
                }

                // Insert the table into the document
                document.InsertTable(table);
            }

            // Save the document
            document.Save();
        }
    }
}
