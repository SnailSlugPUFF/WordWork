using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordWork
{
    class Program
    {
        static void Main(string[] args)
        {
            string startFilePath = "C:\\Users\\User\\Desktop\\doc1.docx";

            string afterFilePath = "C:\\Users\\User\\Desktop\\doc2.docx";

            Dictionary<string, string> tagValues = new Dictionary<string, string>();
            tagValues.Add("[TEXT]", "Привет пока");

            CreateWordDocument(startFilePath, afterFilePath, tagValues);

            Console.WriteLine("Документ успешно создан.");
            Console.ReadLine();
        }

        public static void CreateWordDocument(string startFilePath, string afterFilePath, Dictionary<string, string> tagValues)
        {
            using (WordprocessingDocument word = WordprocessingDocument.Open(startFilePath, true))
            {
                Body body = word.MainDocumentPart.Document.Body;

                foreach (KeyValuePair<string, string> tagValue in tagValues)
                {
                    var tag = new Text(tagValue.Value);
                    var tagElements = body.Descendants<Text>().Where(t => t.Text == tagValue.Key).ToList();
                    foreach (var tagElement in tagElements)
                    {
                        tagElement.Text = "";
                        tagElement.InsertAfterSelf(tag);
                        tagElement.Remove();
                    }
                }

                word.MainDocumentPart.Document.Save();

                word.SaveAs(afterFilePath);
            }
        }
    }
}