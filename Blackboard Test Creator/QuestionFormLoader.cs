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
        public List<OpenXmlElement> QuestionCorrectFeedback { get; set; }
        public List<OpenXmlElement> QuestionIncorrectFeedback { get; set; }
        public List<string> ConstructedQuestionParagraphs { get; set; }
        public List<List<string>> ListOfConstructedQuestionParagraphs { get; set; }
        public List<string> ConstructedAnswerParagraph { get; set; }
        public List<List<string>> ListOfConstructedAnswerParagraphs { get; set; }
        public List<List<List<string>>> ListOfListOfConstructedAnswerParagraphs { get; set; }
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
        public static List<OpenXmlElement> answerPart = new List<OpenXmlElement>();
        List<OpenXmlElement> containerPart = new List<OpenXmlElement>();
        public static List<ImagePart> imgPart1;
        //public static List<ImagePart> imgPart = new List<ImagePart>();
        public static ImagePart[] imgPart;
        public static List<Text> QuestionTopics = new List<Text>();
        public static List<Text> QuestionDifficulty = new List<Text>();
        public static string matchType = string.Empty;
        public static VerticalTextAlignment verticalTextAlignment;
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
                imgPart1 = wordprocessingDocument.MainDocumentPart.ImageParts.ToList();
                imgPart = new ImagePart[imgPart1.Count()];
                foreach (ImagePart part in imgPart1)
                {
                    var resultString = Regex.Match(part.Uri.ToString(), @"\d+").Value;
                    imgPart[Int32.Parse(resultString) - 1] = part;
                }
                Debug.WriteLine(imgPart);
                foreach (OpenXmlElement part in documentParts)
                {
                    if (part.HasAttributes)
                    {
                        foreach (OpenXmlAttribute xmlAttribute in part.GetAttributes())
                        {
                            if(xmlAttribute.Value == "Container")
                            {
                                if (part.Ancestors<SdtContentBlock>().Any())
                                { 
                                    containerPart.Add(part.Ancestors<DocumentFormat.OpenXml.Wordprocessing.SdtContentBlock>().First());
                                }
                                else
                                {
                                    containerPart.Add(part.Parent.Parent.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtContentBlock>().First());
                                }
                            }
                            else if(xmlAttribute.Value == "question")
                            {
                                questionPart.Add(part.Parent.Parent);
                                //questionPart.Add(part.Ancestors<DocumentFormat.OpenXml.Wordprocessing.SdtBlock>().First());
                            }
                            else if(xmlAttribute.Value == "distractor")
                            {
                                answerPart.Add(part.Parent.Parent);
                            }
                            /*
                            if(containerPart.Count() == 0)
                            {
                                containerPart.Add(part.Parent.Parent.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtContentBlock>().First());
                            }
                            */
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
                        if (part.OuterXml.Contains("Type"))
                        {
                            NewQuestion.QuestionType = part.Parent.Parent;
                        }
                        else if (part.OuterXml.Contains("Topics"))
                        {
                            var sdtparent = part.Parent.Parent;
                            foreach (Run run in sdtparent.Descendants<Run>().ToList())
                            {
                                if (run.OuterXml.Contains("PlaceholderText"))
                                {
                                    ;
                                }
                                else
                                {
                                    foreach (Text text in sdtparent.Descendants<Text>().ToList())
                                    {
                                        if (!QuestionTopics.Contains(text))
                                        {
                                            QuestionTopics.Add(text);
                                        }
                                        NewQuestion.Topics.Add(text, NewQuestion.QuestionNumber);
                                    }
                                }
                            }
                        }
                        else if (part.OuterXml.Contains("Level of Difficulty"))
                        {
                            var sdtparent = part.Parent.Parent;
                            foreach (Run run in sdtparent.Descendants<Run>().ToList())
                            {
                                if (run.OuterXml.Contains("PlaceholderText"))
                                {
                                    ;
                                }
                                else
                                {
                                    foreach (Text text in sdtparent.Descendants<Text>().ToList())
                                    {
                                        if (!QuestionDifficulty.Contains(text))
                                        {
                                            QuestionDifficulty.Add(text);
                                        }
                                        NewQuestion.Difficulty.Add(text, NewQuestion.QuestionNumber);
                                    }
                                }
                            }
                        }
                        else if (part.OuterXml.Contains("Match"))
                        {
                            matchType = part.Parent.Parent.InnerText;
                        }
                        else if (part.OuterXml.Contains("distractor"))
                        {
                            if(part.Parent.Parent.Descendants<SdtBlock>().Any())
                            {
                            NewQuestion.AnswerParts = new List<OpenXmlElement>();
                            NewQuestion.AnswerParts = part.Parent.Parent.Descendants<OpenXmlElement>().Last(or => or.Descendants<SdtBlock>().Any()).ToList();
                            }
                        }
                        else if (part.OuterXml.Contains("question feedback correct"))
                        {
                            NewQuestion.QuestionCorrectFeedback = new List<OpenXmlElement>();
                            var sdtparent = part.Parent.Parent;
                            if (sdtparent.InnerText.Contains("(Optional"))
                            {
                                ;
                            }
                            else
                            {
                                NewQuestion.QuestionCorrectFeedback.Add(sdtparent);
                            }
                        }
                        else if (part.OuterXml.Contains("question feedback incorrect"))
                        {
                            NewQuestion.QuestionIncorrectFeedback = new List<OpenXmlElement>();
                            var sdtparent = part.Parent.Parent;
                            if (sdtparent.InnerText.Contains("(Optional"))
                            {
                                ;
                            }
                            else
                            {
                                NewQuestion.QuestionIncorrectFeedback.Add(sdtparent);
                            }
                        }
                    }
                    #region question part
                    NewQuestion.QuestionTextElements = new List<Paragraph>();
                    if (NewQuestion.QuestionItem.Descendants<Paragraph>().Any())
                    {
                       NewQuestion.QuestionTextElements = NewQuestion.QuestionItem.Descendants<Paragraph>().ToList();
                    }
                    else
                    {
                        Paragraph para = new Paragraph();
                        foreach (Run lines in NewQuestion.QuestionItem.Descendants<Run>().ToList())
                        {
                            Run run = para.AppendChild(new Run());
                            if (lines.Descendants<Italic>().Any())
                            {
                                Italic italic = new Italic();
                                run.AppendChild(italic);
                            }
                            if (lines.Descendants<Bold>().Any())
                            {
                                Bold bold = new Bold();
                                run.AppendChild(bold);
                            }
                            if (lines.Descendants<VerticalTextAlignment>().Any())
                            {
                                if (lines.Descendants<VerticalTextAlignment>().First().OuterXml.Contains("superscript"))
                                {
                                    VerticalTextAlignment vertalign = lines.Descendants<VerticalTextAlignment>().First();
                                    verticalTextAlignment = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };
                                }
                                else if (lines.Descendants<VerticalTextAlignment>().First().OuterXml.Contains("subscript"))
                                {
                                    VerticalTextAlignment vertalign = lines.Descendants<VerticalTextAlignment>().First();
                                    verticalTextAlignment = new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript };
                                }
                                run.AppendChild(verticalTextAlignment);
                            }
                            Text text = new Text();
                            text.Text = lines.InnerText;
                            run.AppendChild(text);
                        }
                        NewQuestion.QuestionTextElements.Add(para);          
                    }
                    NewQuestion.ListOfConstructedQuestionParagraphs = new List<List<string>>();
                    foreach (Paragraph paragraph in NewQuestion.QuestionTextElements)
                    {
                        NewQuestion.ConstructedQuestionParagraphs = new List<string>();
                        List<Run> runs = new List<Run>();
                        runs = paragraph.Descendants<Run>().ToList();
                        string runText = string.Empty;
                        foreach (Run run in paragraph.Descendants<Run>().ToList())
                        {
                            runText = run.InnerText;
                            if(run.Descendants<Italic>().Any())
                            {
                                runText = "&lt;i&gt;" + runText + "&lt;/i&gt;";
                            }
                            if (run.Descendants<Bold>().Any())
                            {
                                runText = "&lt;b&gt;" + runText + "&lt;/b&gt;";
                            }
                            if (run.Descendants<Underline>().Any())
                            {
                                runText = "&lt;u&gt;" + runText + "&lt;/u&gt;";
                            }
                            if (run.Descendants<VerticalTextAlignment>().Any())
                            {
                                if (run.Descendants<VerticalTextAlignment>().First().OuterXml.Contains("superscript"))
                                {
                                    runText = "&lt;sup&gt;" + runText + "&lt;/sup&gt;";
                                }
                                else if (run.Descendants<VerticalTextAlignment>().First().OuterXml.Contains("subscript"))
                                {
                                    runText = "&lt;sub&gt;" + runText + "&lt;/sub&gt;";
                                }
                            }
                            NewQuestion.ConstructedQuestionParagraphs.Add(runText);
                        }
                        NewQuestion.ListOfConstructedQuestionParagraphs.Add(NewQuestion.ConstructedQuestionParagraphs);
                    }
                    NewQuestion.QuestionImages = new Dictionary<string, int>();
                    if(NewQuestion.QuestionItem.Descendants<Drawing>().Any())
                    {
                        foreach(Drawing drawing in NewQuestion.QuestionItem.Descendants<Drawing>().AsParallel().ToList())
                        {
                            NewQuestion.QuestionImages.Add("xid-000000" + imageNumber + "_1", questionList.IndexOf(NewQuestion));
                            imageNumber += 1;
                        }
                    }
                    if (NewQuestion.QuestionItem.Descendants<EmbeddedObject>().Any())
                    {
                        foreach (EmbeddedObject embObject in NewQuestion.QuestionItem.Descendants<EmbeddedObject>().AsParallel().ToList())
                        {
                            NewQuestion.QuestionImages.Add("xid-000000" + imageNumber + "_1", questionList.IndexOf(NewQuestion));
                            imageNumber += 1;
                        }
                    }
                    #endregion
                    #region answer part
                    NewQuestion.AnswerParts = new List<OpenXmlElement>();
                    if(NewQuestion.QuestionType.InnerText == "True or False Question")
                    {
                        NewQuestion.AnswerParts.Add(answerPart[i]);
                    }
                    else if(NewQuestion.QuestionType.InnerText == "Essay Question" || NewQuestion.QuestionType.InnerText == "Short Answer Question")
                    {
                        ;
                    }
                    else
                    {
                        NewQuestion.AnswerParts = containerpart.Descendants<OpenXmlElement>().Last(or => or.Descendants<SdtBlock>().Any()).ToList();
                    }
                    NewQuestion.AnswerImages = new Dictionary<string, int>();
                    NewQuestion.ListOfIndividualAnswerParagraphLists = new List<List<Paragraph>>();
                    NewQuestion.ListOfListOfConstructedAnswerParagraphs = new List<List<List<string>>>();
                    foreach(OpenXmlElement answer in NewQuestion.AnswerParts)
                    {
                        NewQuestion.ListOfConstructedAnswerParagraphs = new List<List<string>>();
                        NewQuestion.IndividualAnswerParagraphs = new List<Paragraph>();
                        NewQuestion.IndividualAnswerParagraphs = answer.Descendants<Paragraph>().AsParallel().ToList();
                        NewQuestion.ListOfIndividualAnswerParagraphLists.Add(NewQuestion.IndividualAnswerParagraphs);
                        foreach(Paragraph paragraph in NewQuestion.IndividualAnswerParagraphs)
                        {
                            NewQuestion.ConstructedAnswerParagraph = new List<string>();
                            List<Run> runs = new List<Run>();
                            runs = paragraph.Descendants<Run>().ToList();
                            string runText = string.Empty;
                            foreach(Run run in runs)
                            {
                                runText = run.InnerText;
                                if (run.Descendants<Italic>().Any())
                                {
                                    runText = "&lt;i&gt;" + runText + "&lt;/i&gt;";
                                }
                                if (run.Descendants<Bold>().Any())
                                {
                                    runText = "&lt;b&gt;" + runText + "&lt;/b&gt;";
                                }
                                if (run.Descendants<Underline>().Any())
                                {
                                    runText = "&lt;u&gt;" + runText + "&lt;/u&gt;";
                                }
                                if (run.Descendants<VerticalTextAlignment>().Any())
                                {
                                    if (run.Descendants<VerticalTextAlignment>().First().OuterXml.Contains("superscript"))
                                    {
                                        runText = "&lt;sup&gt;" + runText + "&lt;/sup&gt;";
                                    }
                                    else if (run.Descendants<VerticalTextAlignment>().First().OuterXml.Contains("subscript"))
                                    {
                                        runText = "&lt;sub&gt;" + runText + "&lt;/sub&gt;";
                                    }
                                }
                                NewQuestion.ConstructedAnswerParagraph.Add(runText);
                            }
                            NewQuestion.ListOfConstructedAnswerParagraphs.Add(NewQuestion.ConstructedAnswerParagraph);
                        }
                        NewQuestion.ListOfListOfConstructedAnswerParagraphs.Add(NewQuestion.ListOfConstructedAnswerParagraphs);
                        if (answer.Descendants<Drawing>().Any())
                        {
                            foreach(Drawing drawing in answer.Descendants<Drawing>().AsParallel().ToList())
                            {
                                NewQuestion.AnswerImages.Add("xid-000000" + imageNumber + "_1", NewQuestion.AnswerParts.IndexOf(answer));
                                imageNumber += 1;
                            }
                        }
                        if (answer.Descendants<EmbeddedObject>().Any())
                        {
                            foreach (EmbeddedObject embObject in answer.Descendants<EmbeddedObject>().AsParallel().ToList())
                            {
                                NewQuestion.AnswerImages.Add("xid-000000" + imageNumber + "_1", NewQuestion.AnswerParts.IndexOf(answer));
                                imageNumber += 1;
                            }
                        }
                        if (NewQuestion.IndividualAnswerParagraphs.Count() == 0)
                        {
                            NewQuestion.ListOfIndividualAnswerParagraphLists.Remove(NewQuestion.IndividualAnswerParagraphs);
                        }
                    }
                    #endregion
                    NewQuestion.CorrectAnswers = new List<OpenXmlElement>();
                    foreach (List<Paragraph> list in NewQuestion.ListOfIndividualAnswerParagraphLists)
                    {
                        foreach(OpenXmlElement answer in list)
                        {
                            if(answer.Descendants<Color>().Any())
                            {
                                NewQuestion.CorrectAnswers.Add(answer);
                            }
                            else if(answer.Descendants<Highlight>().Any())
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
