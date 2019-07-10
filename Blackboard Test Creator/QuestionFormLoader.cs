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
        public int QuestionNumber { get; set; }
        public List<DocumentFormat.OpenXml.Wordprocessing.Text> QuestionTextElements { get; set; }
    }
    class QuestionFormLoader
    {
        public string formPath = Form1.TestFormFilePath;
        public static Word.Application wordApp = new Word.Application();
        public static Word.Document Form;
        public static WordprocessingDocument wordprocessingDocument;
        public static XmlDocument XMLForm = new XmlDocument();
        List<DocumentFormat.OpenXml.Wordprocessing.Text> contentBlockParts;
        //OpenXmlElement part2ndchild;
        //OpenXmlElement part3rdchild;
        List<OpenXmlElement> containerPart = new List<OpenXmlElement>();
        public void FormLoader()
        {
            if (formPath != null)
            {
                //Form = wordApp.Documents.Open(formPath);
                //string content = GetWordDocumentContent(formPath);
                Stream stream = File.Open(formPath, FileMode.Open);
                wordprocessingDocument = WordprocessingDocument.Open(stream, true);
                List<OpenXmlElement> documentParts = new List<OpenXmlElement>();
                List<DocumentFormat.OpenXml.OpenXmlAttribute> partAttributes = new List<OpenXmlAttribute>();
                documentParts = wordprocessingDocument.MainDocumentPart.Document.Body.Descendants().ToList();
                foreach (OpenXmlElement part in documentParts)
                {
                    if(part.HasAttributes)
                    {
                        foreach (OpenXmlAttribute xmlAttribute in part.GetAttributes())
                        {
                            if(xmlAttribute.Value == "Container")
                            {
                                Debug.WriteLine("container part");
                                Debug.WriteLine(part.Ancestors<DocumentFormat.OpenXml.Wordprocessing.SdtBlock>().First().InnerText);
                                containerPart.Add(part.Ancestors<DocumentFormat.OpenXml.Wordprocessing.SdtBlock>().First());
                            }
                        }
                    }
                    /*
                    var partType = part.GetType().ToString();
                    if(partType == "DocumentFormat.OpenXml.Wordprocessing.SdtContentBlock")
                    {
                        var part1stchild = part.FirstChild;
                        if (part1stchild.HasChildren) {part2ndchild = part1stchild.FirstChild;}
                        if (part2ndchild.HasChildren) { part3rdchild = part2ndchild.FirstChild; }
                        if(part3rdchild.ToString() == "DocumentFormat.OpenXml.Wordprocessing.Tag")
                        {
                            contentBlockParts = new List<OpenXmlElement>();
                            contentBlockParts = part.Descendants().ToList();
                            foreach (OpenXmlElement openXmlElement in contentBlockParts)
                            {
                            }
                        }
                        */
                    /*
                    if (part.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.SdtBlock>().GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.SdtProperties>().GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Tag>().ToString() != null)
                    {
                        var questionBlock = part.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.SdtBlock>().GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.SdtProperties>().GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Tag>();
                        Debug.WriteLine(questionBlock);
                    }
                    */
                    /*
                    if (questionBlock.ToString() == "DocumentFormat.OpenXml.Wordprocessing.Tag" && questionBlock != null)
                    {
                        contentBlockParts = new List<OpenXmlElement>();
                        contentBlockParts = part.ChildElements.ToList();
                        foreach (OpenXmlElement openXmlElement in contentBlockParts)
                        {
                            Debug.WriteLine(openXmlElement);
                        }
                    }

                }
                */
                    //Debug.WriteLine(part);
                }
                foreach(OpenXmlElement containerpart in containerPart)
                {
                    Question NewQuestion = new Question();
                    NewQuestion.QuestionItem = containerpart;
                    NewQuestion.QuestionTextElements = new List<Text>();
                    NewQuestion.QuestionTextElements = containerpart.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList();
                    foreach (Text text in NewQuestion.QuestionTextElements)
                    {
                        Debug.WriteLine(text.InnerText);
                    }
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
