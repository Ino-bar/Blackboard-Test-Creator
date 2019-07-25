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
        public OpenXmlElement QuestionType { get; set; }
        public int QuestionNumber { get; set; }
        public List<Paragraph> QuestionTextElements { get; set; }
        public List<OpenXmlElement> AnswerParts { get; set; }
        public List<Paragraph> IndividualAnswerParagraphs { get; set; }
        public List<List<Paragraph>> ListOfIndividualAnswerParagraphLists { get; set; }
        public List<OpenXmlElement> CorrectAnswers { get; set; }
        public Dictionary<string, int> QuestionImages { get; set; }
        public Dictionary<string, int> AnswerImages { get; set; }
        public Dictionary<Text, int> Topics { get; set; }
        public Dictionary<Text, int> Difficulty { get; set; }
    }
    class QuestionFormLoader
    {
        public string formPath = Form1.TestFormFilePath;
        //public static Word.Application wordApp = new Word.Application();
        public static Stream stream;
        public static WordprocessingDocument wordprocessingDocument;
        public static XmlDocument XMLForm = new XmlDocument();
        public static List<Question> questionList = new List<Question>();
        public static List<OpenXmlElement> questionPart = new List<OpenXmlElement>();
        List<OpenXmlElement> containerPart = new List<OpenXmlElement>();
        public static List<ImagePart> imgPart;
        public static List<Text> QuestionTopics = new List<Text>();
        public static List<Text> QuestionDifficulty = new List<Text>();
        int imageNumber = 1;
        public void FormLoader()
        {
            if (formPath != null)
            {
                stream = File.Open(formPath, FileMode.Open);
                wordprocessingDocument = WordprocessingDocument.Open(stream, true);
                List<OpenXmlElement> documentParts = new List<OpenXmlElement>();
                List<DocumentFormat.OpenXml.OpenXmlAttribute> partAttributes = new List<OpenXmlAttribute>();
                documentParts = wordprocessingDocument.MainDocumentPart.Document.Body.Descendants().ToList();
                imgPart = wordprocessingDocument.MainDocumentPart.ImageParts.ToList();
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
                    NewQuestion.QuestionItem = questionPart[i];
                    NewQuestion.QuestionNumber = i + 1;
                    NewQuestion.Topics = new Dictionary<Text, int>();
                    NewQuestion.Difficulty = new Dictionary<Text, int>();
                    foreach (Tag part in containerpart.Descendants<Tag>())
                    { 
                        if(part.OuterXml.Contains("Type"))
                        {
                            NewQuestion.QuestionType = part.Parent.Parent;
                        }
                        else if(part.OuterXml.Contains("Topics"))
                        {
                            var sdtparent = part.Parent.Parent;
                            foreach(Text text in sdtparent.Descendants<Text>().ToList())
                            {
                                if (!QuestionTopics.Contains(text))
                                { 
                                    QuestionTopics.Add(text);
                                }
                                NewQuestion.Topics.Add(text, NewQuestion.QuestionNumber);
                            }
                        }
                        else if(part.OuterXml.Contains("Level of Difficulty"))
                        {
                            var sdtparent = part.Parent.Parent;
                            foreach (Text text in sdtparent.Descendants<Text>().ToList())
                            {
                                if(!QuestionDifficulty.Contains(text))
                                {
                                    QuestionDifficulty.Add(text);
                                }
                                NewQuestion.Difficulty.Add(text, NewQuestion.QuestionNumber);
                            }
                        }
                    }
                    NewQuestion.QuestionTextElements = new List<Paragraph>();
                    NewQuestion.QuestionTextElements = NewQuestion.QuestionItem.Descendants<Paragraph>().ToList();
                    NewQuestion.QuestionImages = new Dictionary<string, int>();
                    if(NewQuestion.QuestionItem.Descendants<Drawing>().Any())
                    {
                        foreach(Drawing drawing in NewQuestion.QuestionItem.Descendants<Drawing>().AsParallel().ToList())
                        {
                            NewQuestion.QuestionImages.Add("xid-000000" + imageNumber + "_1", questionList.IndexOf(NewQuestion));
                            imageNumber += 1;
                        }
                    }
                    NewQuestion.AnswerParts = new List<OpenXmlElement>();
                    NewQuestion.AnswerParts = containerpart.Descendants<OpenXmlElement>().Last(or => or.Descendants<SdtBlock>().Any()).ToList();
                    NewQuestion.AnswerImages = new Dictionary<string, int>();
                    NewQuestion.ListOfIndividualAnswerParagraphLists = new List<List<Paragraph>>();
                    foreach(OpenXmlElement answer in NewQuestion.AnswerParts)
                    {
                        NewQuestion.IndividualAnswerParagraphs = new List<Paragraph>();
                        NewQuestion.IndividualAnswerParagraphs = answer.Descendants<Paragraph>().AsParallel().ToList();
                        NewQuestion.ListOfIndividualAnswerParagraphLists.Add(NewQuestion.IndividualAnswerParagraphs);
                        if(answer.Descendants<Drawing>().Any())
                        {
                            foreach(Drawing drawing in answer.Descendants<Drawing>().AsParallel().ToList())
                            {
                                NewQuestion.AnswerImages.Add("xid-000000" + imageNumber + "_1", NewQuestion.AnswerParts.IndexOf(answer));
                                imageNumber += 1;
                            }
                        }
                    }
                    NewQuestion.CorrectAnswers = new List<OpenXmlElement>();
                    foreach (List<Paragraph> list in NewQuestion.ListOfIndividualAnswerParagraphLists)
                    {
                        foreach(OpenXmlElement answer in list)
                        {
                            if(answer.Descendants<Color>().Any())
                            {
                                NewQuestion.CorrectAnswers.Add(answer);
                            }
                        }
                    }
                    i++;
                }
            }
        }
    }
}
