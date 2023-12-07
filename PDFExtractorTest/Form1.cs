using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Org.BouncyCastle.Asn1.Ocsp;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static iTextSharp.text.pdf.codec.TiffWriter;
using Excel = Microsoft.Office.Interop.Excel;

namespace PDFExtractorTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            string pdfPath = "C:\\temp_nci\\Credit_Application.pdf";
            string exclusionFilePath = ".\\exclusion_file.txt"; // Path to the file containing phrases to exclude from processing
            string fieldsListPath = ".\\fields_list.json"; // Path to the file containing field names

            string[] exclusionPhrases = File.ReadAllLines(exclusionFilePath);
            HashSet<string> exclusions = new HashSet<string>(exclusionPhrases);

            // Load field names from JSON file
            string jsonText = File.ReadAllText(fieldsListPath);
            var fieldCategories = JsonConvert.DeserializeObject<Dictionary<string, JToken>>(jsonText);

            // Combine all fields into a single list
            List<string> fieldNames = new List<string>();

            foreach (var category in fieldCategories)
            {
                AddFieldNamesFromToken(category.Value, fieldNames);
            }

            PdfReader reader = new PdfReader(pdfPath);

            StringBuilder text = new StringBuilder();

            // Step 2: Process each page
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                string pageText = PdfTextExtractor.GetTextFromPage(reader, i);

                string[] lines = pageText
                    .Replace("\r", "")
                    .Replace('\n', ',')
                    .Replace(" Postcode ", ",Postcode ")
                    .Replace(" Mobile ", ",Mobile ")
                    .Split(',');

                for (var line = 0; line < lines.Length; line++)
                {
                    bool excludeLine = false;

                    foreach (string phrase in exclusions)
                    {
                        if (lines[line].StartsWith(phrase))
                        {
                            excludeLine = true;
                            break;
                        }
                    }
                    if (lines[line].StartsWith("How long have you been trading in current"))
                    {
                        MergeSpecificLines(lines);
                    }
                    
                    
                    // Check if the line contains any of the field names
                    bool containsFieldName = fieldNames.Any(fieldName =>
                    {
                        bool result = lines[line].StartsWith(fieldName);
                        return result;
                    });


                    if (lines[line].StartsWith("Address2"))
                    {
                        containsFieldName = true;
                    }

                    bool isJoined = false;

                    foreach (string keyField in fieldNames)
                    {
                        bool isLinesNkeyField = lines[line].Equals(keyField);
                        bool isNextAFieldName = !fieldNames.Any(fieldName =>
                        {
                            if (line < lines.Length - 1)
                            {
                                return lines[line + 1].StartsWith(fieldName);
                            }
                            else
                            {
                                return false;
                            }
                        });

                        if (isLinesNkeyField && isNextAFieldName)

                        {
                            isJoined = true;
                            break;
                        }
                    }

                    if (!excludeLine
                        && containsFieldName
                        && lines[line].Trim().Length > 0
                        && line < lines.Length)
                    {
                        if (isJoined)
                        {
                            // if current line is supposed to be joined to next line
                            // then temp store current line and increment lines
                            string currentLine = lines[line];
                            text.AppendLine(currentLine + " " + lines[++line]);
                            line--;

                            isJoined = false;

                        }
                        else
                        { // handling 2 fields for each address group
                            if (
                                lines[line].StartsWith("Address")
                                || lines[line].StartsWith("Street address")
                                || lines[line].StartsWith("Delivery address")
                                )
                            {
                                text.AppendLine(lines[line]);

                                text.AppendLine("Address2 " + lines[++line]);
                                line--;

                            }
                            else
                            { // normal default appending
                                text.AppendLine(lines[line]);

                            }
                        }
                        excludeLine = false;
                        containsFieldName = false;
                    }
                }

            }
            text.Replace("Preferred order method\r\nPhone order", "Preferred order method Phone order");
            reader.Close();

            // Step 3: Extract field values
            string allText = text.ToString();

            System.Console.Write(text);

            // Assuming allLines is an array of strings, each representing a line of text
            string[] allLines = allText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            //allLines = MergeSpecificLines(allLines); // Adjust this method to return string[]

            List<KeyValuePair<string, string>> fieldValues = new List<KeyValuePair<string, string>>();

            foreach (var line in allLines)
            {
                foreach (var fieldName in fieldNames)
                {
                    if (line.StartsWith(fieldName))
                    {
                        // Assuming the format is "fieldName: value"
                        var fieldValue = line.Substring(fieldName.Length).Trim();

                        // Add the field name and value as a key-value pair
                        fieldValues.Add(new KeyValuePair<string, string>(fieldName, fieldValue));

                        break; // Found the field in this line, no need to check other field names
                    }
                }
            }

            sendToExcel(fieldValues);
        }

        private static void AddFieldNamesFromToken(JToken token, List<string> fieldNames)
        {
            switch (token.Type)
            {
                case JTokenType.Array:
                    foreach (var item in token.Children())
                    {
                        AddFieldNamesFromToken(item, fieldNames);
                    }
                    break;
                case JTokenType.String:
                    fieldNames.Add(token.ToString());
                    break;
                    // Add cases for other JToken types if necessary
            }
        }

        private static string[] MergeSpecificLines(string[] input)
        {
            // Define the split line and how it should be merged
            string splitLineStart = "How long have you been trading in current";
            string splitLineEnd = "location?";

            for (int i = 0; i < input.Length - 1; i++)
            {
                if (input[i].StartsWith(splitLineStart) && input[i + 1].Equals(splitLineEnd))
                {
                    string fieldValue = 
                        input[i].Substring(splitLineStart.Length, 
                        input[i].Length - splitLineStart.Length).Trim();
                    
                    input[i] =
                        splitLineStart
                        + " "
                        + splitLineEnd
                        + " "
                        + fieldValue;

                    input[i + 1] = ""; // Optionally, you could remove or leave empty the next line
                    break; // Assuming only one occurrence. Remove break if multiple occurrences are possible.
                }
            }

            return input;
        }


        private static void sendToExcel(List<KeyValuePair<string, string>> fieldValues)
        {
            // Opening the Excel application and workbook
            Excel.Application excelApp = new Excel.Application();

            // todo: change to false for production
            excelApp.Visible = true;
            string workbookPath = @"C:\Users\infil\Documents\NCI_files\nci_credit_applications_database.xlsm";

            DateTime now = DateTime.Now;
            string workbookPathNew = $"C:\\Users\\infil\\Documents\\NCI_files\\nci_credit_applications_database_{now.ToString("yyyyMMdd_HHmmss")}.xlsm";
            Excel.Workbook workbook = excelApp.Workbooks.Open(workbookPath);
            Excel.Worksheet worksheet = workbook.Sheets["All Fields"];

            // Find the first empty row
            int row = 1;
            while (worksheet.Cells[row, 1].Value2 != null)
            {
                row++;
            }

            int col = 1; // Starting column (A)

            foreach (var entry in fieldValues)
            {
                string key = entry.Key;
                string value = entry.Value;

                // Skip the headings and insert only the contents
                if (!key.StartsWith("The Applicant") &&
                    !key.StartsWith("Company details") &&
                    !key.StartsWith("Director details") &&
                    !key.StartsWith("Contact details") &&
                    !key.StartsWith("Billing address") &&
                    !key.StartsWith("Delivery address for stock") &&
                    !key.StartsWith("Email address") &&
                    !key.StartsWith("Trading details") &&
                    !key.StartsWith("Authorised Signatory"))
                {
                    // Ensure we don't go beyond column BL
                    if (col > 72) break;

                    //worksheet.Cells[row, col].Value2 = key;
                    worksheet.Cells[row, col].Value2 = value;
                    col++; // Move two columns to the right for each entry
                }
            }

            // Save and close the workbook
            workbook.SaveAs(workbookPathNew);
            workbook.Close();
            excelApp.Quit();

            // Release the COM objects
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

    }
}
