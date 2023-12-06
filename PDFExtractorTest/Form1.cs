using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using static System.Net.Mime.MediaTypeNames;

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
                string[] lines = pageText.Split('\n');

                foreach (string line in lines)
                {
                    bool excludeLine = false;

                    foreach (string phrase in exclusions)
                    {
                        if (line.Contains(phrase))
                        {
                            excludeLine = true;
                            break;
                        }
                    }

                    if (!excludeLine && line.Trim().Length > 0)
                    {
                        text.AppendLine(line); // Append the line if it does not contain any exclusion phrases
                        rtbStats.AppendText(line + '\n');
                    }
                }

            }




            reader.Close();

            // Step 3: Extract field values
            string allText = text.ToString();

            // Merge specific split lines
            allText = MergeSpecificLines(allText);
            Dictionary<string, string> fieldValues = new Dictionary<string, string>();
            for (int i = 0; i < fieldNames.Count; i++)
            {
                string startField = fieldNames[i];
                string endField = i < fieldNames.Count - 1 ? fieldNames[i + 1] : null;
                int startIndex = allText.IndexOf(startField) + startField.Length;
                int endIndex = endField != null ? allText.IndexOf(endField, startIndex) : allText.Length;

                string fieldValue = startIndex < endIndex ? allText.Substring(startIndex, endIndex - startIndex).Trim() : "";
                fieldValues[startField] = fieldValue;
            }

            // Output the field values
            foreach (var fieldValue in fieldValues)
            {
                Console.WriteLine($"{fieldValue.Key}: {fieldValue.Value}");
            }

        }
        private static IEnumerable<string> FlattenFieldList(JToken token)
        {
            if (token.Type == JTokenType.Array)
            {
                foreach (var item in token.Children())
                {
                    if (item.Type == JTokenType.String)
                    {
                        yield return item.ToString();
                    }
                    else if (item.Type == JTokenType.Array)
                    {
                        foreach (var nestedItem in item.Children())
                        {
                            if (nestedItem.Type == JTokenType.String)
                            {
                                yield return nestedItem.ToString();
                            }
                        }
                    }
                }
            }
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

        private static string MergeSpecificLines(string text)
        {
            // Define the split line and how it should be merged
            string splitLineStart = "How long have you been trading in current";
            string splitLineEnd = "location?";

            // Replace the newline between the split lines with a space
            return text.Replace(splitLineStart + "\n" + splitLineEnd, splitLineStart + " " + splitLineEnd);
        }

    }
}
