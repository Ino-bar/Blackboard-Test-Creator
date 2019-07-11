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
    public class Question
    {
        public OpenXmlElement QuestionItem { get; set; }
        public OpenXmlElement QuestionNumber { get; set; }
        public List<DocumentFormat.OpenXml.Wordprocessing.Text> QuestionTextElements { get; set; }
        public List<OpenXmlElement> AnswerParts { get; set; }
    }
    class QuestionFormLoader
    {
        public string formPath = Form1.TestFormFilePath;
        //public static Word.Application wordApp = new Word.Application();
        public static Stream stream;
        public static Word.Document Form;
        public static WordprocessingDocument wordprocessingDocument;
        public static XmlDocument XMLForm = new XmlDocument();
        public static List<Question> questionList = new List<Question>();
        public static List<OpenXmlElement> questionPart = new List<OpenXmlElement>();
        List<OpenXmlElement> containerPart = new List<OpenXmlElement>();
        public void FormLoader()
        {
            if (formPath != null)
            {
                //Form = wordApp.Documents.Open(formPath);
                //string content = GetWordDocumentContent(formPath);
                stream = File.Open(formPath, FileMode.Open);
                wordprocessingDocument = WordprocessingDocument.Open(stream, true);
                List<OpenXmlElement> documentParts = new List<OpenXmlElement>();
                List<DocumentFormat.OpenXml.OpenXmlAttribute> partAttributes = new List<OpenXmlAttribute>();
                documentParts = wordprocessingDocument.MainDocumentPart.Document.Body.Descendants().ToList();
                foreach (OpenXmlElement part in documentParts)
                {
                    if (part.HasAttributes)
                    {
                        foreach (OpenXmlAttribute xmlAttribute in part.GetAttributes())
                        {
                            if(xmlAttribute.Value == "Container")
                            {
                                containerPart.Add(part.Ancestors<DocumentFormat.OpenXml.Wordprocessing.SdtContentBlock>().First());
                            }
                            else if(xmlAttribute.Value == "question")
                            {
                                questionPart.Add(part.Parent.Parent);
                            }
                        }
                    }
                }
                var i = 0;
                foreach(OpenXmlElement containerpart in containerPart)
                {
                    Question NewQuestion = new Question();
                    questionList.Add(NewQuestion);
                    NewQuestion.AnswerParts = new List<OpenXmlElement>();
                    NewQuestion.AnswerParts = containerpart.Descendants<OpenXmlElement>().Last(or => or.Descendants<SdtBlock>().Any()).ToList();
                    NewQuestion.QuestionItem = containerpart;
                    NewQuestion.QuestionNumber = questionPart[i];
                    NewQuestion.QuestionTextElements = new List<Text>();
                    NewQuestion.QuestionTextElements = containerpart.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList();
                    Debug.WriteLine(NewQuestion.QuestionNumber.InnerText);
                    foreach (OpenXmlElement element in NewQuestion.AnswerParts)
                    {
                        Debug.WriteLine(element.InnerText);
                        if(element.InnerXml.Contains("FF0000"))
                        {
                            Debug.WriteLine("Answer " + (NewQuestion.AnswerParts.IndexOf(element) + 1) + " is correct");
                        }
                    }
                    /*
                    foreach (Text text in NewQuestion.QuestionTextElements)
                    {
                        Debug.WriteLine(text.InnerText);
                    }
                    */
                    i++;
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
