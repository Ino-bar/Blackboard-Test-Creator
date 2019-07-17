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
using System.Drawing;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace Blackboard_Test_Creator
{
    public class Question
    {
        public OpenXmlElement QuestionItem { get; set; }
        public OpenXmlElement QuestionNumber { get; set; }
        public List<Paragraph> QuestionTextElements { get; set; }
        public List<OpenXmlElement> AnswerParts { get; set; }
        public List<Paragraph> IndividualAnswerParagraphs { get; set; }
        public List<List<Paragraph>> ListOfIndividualAnswerParagraphLists { get; set; }
        public List<OpenXmlElement> CorrectAnswers { get; set; }
        public List<ImagePart> AnswerImages { get; set; }
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
        public static List<ImagePart> imgPart;
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
                imgPart = wordprocessingDocument.MainDocumentPart.ImageParts.ToList();
                foreach (OpenXmlElement part in documentParts)
                {
                    //Debug.WriteLine(part);
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
                                questionPart.Add(part.Ancestors<DocumentFormat.OpenXml.Wordprocessing.SdtBlock>().First());
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
                    NewQuestion.ListOfIndividualAnswerParagraphLists = new List<List<Paragraph>>();
                    foreach(OpenXmlElement answer in NewQuestion.AnswerParts)
                    {
                        NewQuestion.IndividualAnswerParagraphs = new List<Paragraph>();
                        NewQuestion.IndividualAnswerParagraphs = answer.Descendants<Paragraph>().AsParallel().ToList();
                        NewQuestion.ListOfIndividualAnswerParagraphLists.Add(NewQuestion.IndividualAnswerParagraphs);
                    }
                    NewQuestion.QuestionItem = questionPart[i];
                    NewQuestion.QuestionNumber = questionPart[i];
                    NewQuestion.QuestionTextElements = new List<Paragraph>();
                    NewQuestion.QuestionTextElements = NewQuestion.QuestionItem.Descendants<Paragraph>().ToList();
                    NewQuestion.CorrectAnswers = new List<OpenXmlElement>();
                    foreach (Paragraph questiontext in NewQuestion.QuestionTextElements)
                    { 
                        Debug.WriteLine(questiontext.InnerText);
                    }
                    foreach (List<Paragraph> list in NewQuestion.ListOfIndividualAnswerParagraphLists)
                    {
                        foreach(OpenXmlElement answer in list)
                        {
                            Debug.WriteLine(answer.InnerText);
                            if(answer.Descendants<Color>().Any())
                            {
                                Debug.WriteLine("Answer " + (NewQuestion.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1) + " is correct");
                                NewQuestion.CorrectAnswers.Add(answer);
                            }
                        }
                    }
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
