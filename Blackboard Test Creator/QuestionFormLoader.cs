using System;
using System.Collections.Generic;
using System.Linq;
using System.IO.Packaging;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.ComponentModel;
using System.Xml.Linq;
using System.Xml;
using Microsoft.Office.Tools.Word;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office.Word;
using System.Windows;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Blackboard_Test_Creator
{
    class QuestionFormLoader
    {
        public string formPath = Form1.TestFormFilePath;
        public static Word.Application wordApp = new Word.Application();
        public static Word.Document Form;
        public static WordprocessingDocument wordprocessingDocument;
        public static XmlDocument XMLForm = new XmlDocument();
        public void FormLoader()
        {
            if (formPath != null)
            {
                //Form = wordApp.Documents.Open(formPath);
                //string content = GetWordDocumentContent(formPath);
                //Debug.WriteLine(content);
                Stream stream = File.Open(formPath, FileMode.Open);
                wordprocessingDocument = WordprocessingDocument.Open(stream, true);
                List<OpenXmlElement> documentParts = new List<OpenXmlElement>();
                
                documentParts = wordprocessingDocument.MainDocumentPart.Document.Body.Descendants().ToList();
                foreach (OpenXmlElement part in documentParts)
                {
                    documentParts.Add(part);
                    Debug.WriteLine(part);
                }
            }
        }
        private static string GetWordDocumentContent(string strDoc)
        {
            Stream stream = File.Open(strDoc, FileMode.Open);
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true);
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            string content = body.InnerText;
            return content;
        }
    }
}
