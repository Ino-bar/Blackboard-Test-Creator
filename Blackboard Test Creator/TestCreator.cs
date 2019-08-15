using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;
using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using System.IO.Compression;


namespace Blackboard_Test_Creator
{
    class TestCreator
    {
        int imagenumber = 1;
        public string savePath = Form1.TestFilePath;
        string questionType = string.Empty;
        string negativepointsind = string.Empty;
        string rcardinality = string.Empty;
        string answerResult = string.Empty;
        string correctAnswer = string.Empty;
        static int totalScore = QuestionFormLoader.questionList.Count() * Form1.QuestionScore;
        List<string> questionParagraphs = new List<string>();
        List<string> answerParagraphs = new List<string>();
        Dictionary<string, string> fileList = new Dictionary<string, string>();
        string[] res0001assessdata =
        {
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
            "<questestinterop>",
            "<assessment title=\"" + Form1.TestName + "\">",
            "<assessmentmetadata>",
            "<bbmd_asi_object_id>" + Form1.TestName + " " + DateTime.Today + "</bbmd_asi_object_id>",
            "<bbmd_asitype>Assessment</bbmd_asitype>",
            "<bbmd_assessmenttype>Test</bbmd_assessmenttype>",
            "<bbmd_sectiontype>Subsection</bbmd_sectiontype>",
            "<bbmd_questiontype>Multiple Choice</bbmd_questiontype>",
            "<bbmd_is_from_cartridge>false</bbmd_is_from_cartridge>",
            "<bbmd_is_disabled>false</bbmd_is_disabled>",
            "<bbmd_negative_points_ind>N</bbmd_negative_points_ind>",
            "<bbmd_canvas_fullcrdt_ind>false</bbmd_canvas_fullcrdt_ind>",
            "<bbmd_all_fullcredit_ind>false</bbmd_all_fullcredit_ind>",
            "<bbmd_numbertype>none</bbmd_numbertype>",
            "<bbmd_partialcredit>",
            "</bbmd_partialcredit>",
            "<bbmd_orientationtype>vertical</bbmd_orientationtype>",
            "<bbmd_is_extracredit>false</bbmd_is_extracredit>",
            "<qmd_absolutescore_max>" + totalScore + "</qmd_absolutescore_max>",
            "<qmd_weighting>0.0</qmd_weighting>",
            "<qmd_instructornotes>",
            "</qmd_instructornotes>",
            "</assessmentmetadata>",
            "<rubric view=\"All\">",
            "<flow_mat class=\"Block\">",
            "<material>",
            "<mat_extension>",
            "<mat_formattedtext type=\"HTML\" />",
            "</mat_extension>",
            "</material>",
            "</flow_mat>",
            "</rubric>",
            "<presentation_material>",
            "<flow_mat class=\"Block\">",
            "<material>",
            "<mat_extension>",
            "<mat_formattedtext type=\"HTML\" />",
            "</mat_extension>",
            "</material>",
            "</flow_mat>",
            "</presentation_material>",
            "<section>",
            "<sectionmetadata>",
            "<bbmd_asi_object_id>section_0</bbmd_asi_object_id>",
            "<bbmd_asitype>Section</bbmd_asitype>",
            "<bbmd_assessmenttype>Test</bbmd_assessmenttype>",
            "<bbmd_sectiontype>Subsection</bbmd_sectiontype>",
            "<bbmd_questiontype>Multiple Choice</bbmd_questiontype>",
            "<bbmd_is_from_cartridge>false</bbmd_is_from_cartridge>",
            "<bbmd_is_disabled>false</bbmd_is_disabled>",
            "<bbmd_negative_points_ind>N</bbmd_negative_points_ind>",
            "<bbmd_canvas_fullcrdt_ind>false</bbmd_canvas_fullcrdt_ind>",
            "<bbmd_all_fullcredit_ind>false</bbmd_all_fullcredit_ind>",
            "<bbmd_numbertype>none</bbmd_numbertype>",
            "<bbmd_partialcredit>",
            "</bbmd_partialcredit>",
            "<bbmd_orientationtype>vertical</bbmd_orientationtype>",
            "<bbmd_is_extracredit>false</bbmd_is_extracredit>",
            "<qmd_absolutescore_max>" + totalScore + "</qmd_absolutescore_max>",
            "<qmd_weighting>0.0</qmd_weighting>",
            "<qmd_instructornotes>",
            "</qmd_instructornotes>",
            "</sectionmetadata>"
        };
        string[] res0001assessdataend =
        {
            "</section>",
            "</assessment>",
            "</questestinterop>"
        };
        public FileStream CreateFile(string FilePath, string FileName)
        {
            FileStream file = File.Create(FilePath + "\\" + FileName);
            return file;
        }
        public void CreatecsfilesFolder()
        {
            Directory.CreateDirectory(savePath + "\\csfiles\\home_dir");
        }
        public void SaveImageFiles()
        {
            int i = 1;

            foreach (ImagePart image in QuestionFormLoader.imgPart)
            {
                string[] lines =
                {
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                    "<lom xmlns=\"http://www.imsglobal.org/xsd/imsmd_rootv1p2p1\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:schemaLocation=\"http://www.imsglobal.org/xsd/imsmd_rootv1p2p1 imsmd_rootv1p2p1.xsd\">",
                    "<relation>",
                    "<resource>",
                    "<identifier>000000" + i + "_1#/courses/FAKE-COURSE/" + Form1.TestName + "/image00" + i + ".png</identifier>",
                    "</resource>",
                    "</relation>",
                    "</lom>"
                };
                Image img = Image.FromStream(image.GetStream());
                string fileName = savePath + "\\csfiles\\home_dir\\image00" + i + "__xid-000000" + i + "_1.png";
                img.Save(fileName);
                string path = savePath + "\\csfiles\\home_dir\\image00" + i + "__xid-000000" + i + "_1.png.xml";
                using (StreamWriter imgxmlfile = new StreamWriter(path))
                {
                    foreach (string line in lines)
                        imgxmlfile.WriteLine(line);
                }
                i++;
            }
        }
        public void CreateBBPackage()
        {
            string path = savePath + "\\.bb-package-info";
            using (StreamWriter BBPackage = new StreamWriter(path))
            {
                BBPackage.WriteLine("cx.package.info.version=6.0");
            }
            fileList.Add(path, ".bb-package-info");
        }
        public void Createimsmanifest()
        {
            if (QuestionFormLoader.QuestionTopics.Count() == 0 && QuestionFormLoader.QuestionDifficulty.Count() == 0)
            { 
                string[] manifestwithouttopics =
                {
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                    "<manifest identifier=\"man00001\"",
                    "xmlns:bb=\"http://www.blackboard.com/content-packaging/\">",
                    "<organizations default=\"toc00001\">",
                    "<organization identifier=\"toc00001\" />",
                    "</organizations>",
                    "<resources>",
                    "<resource",
                        "bb:file=\"res00001.dat\"",
                        "bb:title=\"" + Form1.TestName + "\"",
                        "xml:base=\"res00001\"",
                        "identifier=\"res00001\"",
                        "type=\"assessment/x-bb-qti-test\" />",
                    "<resource",
                        "bb:file=\"res00002.dat\"",
                        "bb:title=\"Assessment Creation Settings\"",
                        "xml:base=\"res00002\"",
                        "identifier=\"res00002\"",
                        "type=\"course/x-bb-courseassessmentcreationsettings\" />",
                    "<resource",
                        "bb:file=\"res00003.dat\"",
                        "bb:title=\"LearnRubrics\"",
                        "xml:base=\"res00003\"",
                        "identifier=\"res00003\"",
                        "type=\"course/x-bb-rubrics\" />",
                    "<resource",
                        "bb:file=\"res00004.dat\"",
                        "bb:title=\"CSResourceLinks\"",
                        "xml:base=\"res00004\"",
                        "identifier=\"res00004\"",
                        "type=\"course/x-bb-csresourcelinks\" />",
                    "<resource",
                        "bb:file=\"res00005.dat\"",
                        "bb:title=\"CourseRubricAssociation\"",
                        "xml:base=\"res00005\"",
                        "identifier=\"res00005\"",
                        "type=\"course/x-bb-crsrubricassocation\" />",
                    "</resources>",
                    "</manifest>"
                };
                string path = savePath + "\\imsmanifest.xml";
                using (StreamWriter imsmanifest = new StreamWriter(path))
                {
                    foreach (string line in manifestwithouttopics)
                        imsmanifest.WriteLine(line);
                }
                fileList.Add(path, "imsmanifest.xml");
            }
            else { 
                string[] manifestwithtopics =
                {
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                    "<manifest identifier=\"man00001\"",
                    "xmlns:bb=\"http://www.blackboard.com/content-packaging/\">",
                    "<organizations default=\"toc00001\">",
                    "<organization identifier=\"toc00001\" />",
                    "</organizations>",
                    "<resources>",
                    "<resource",
                        "bb:file=\"res00001.dat\"",
                        "bb:title=\"" + Form1.TestName + "\"",
                        "xml:base=\"res00001\"",
                        "identifier=\"res00001\"",
                        "type=\"assessment/x-bb-qti-test\" />",
                    "<resource",
                        "bb:file=\"res00002.dat\"",
                        "bb:title=\"Categories\"",
                        "xml:base=\"res00002\"",
                        "identifier=\"res00002\"",
                        "type=\"course/x-bb-category\" />",
                    "<resource",
                        "bb:file=\"res00003.dat\"",
                        "bb:title=\"Item Categories\"",
                        "xml:base=\"res00003\"",
                        "identifier=\"res00003\"",
                        "type=\"course/x-bb-itemcategory\" />",
                    "<resource",
                        "bb:file=\"res00004.dat\"",
                        "bb:title=\"Assessment Creation Settings\"",
                        "xml:base=\"res00004\"",
                        "identifier=\"res00004\"",
                        "type=\"course/x-bb-courseassessmentcreationsettings\" />",
                    "<resource",
                        "bb:file=\"res00005.dat\"",
                        "bb:title=\"LearnRubrics\"",
                        "xml:base=\"res00005\"",
                        "identifier=\"res00005\"",
                        "type=\"course/x-bb-rubrics\" />",
                    "<resource",
                        "bb:file=\"res00006.dat\"",
                        "bb:title=\"CSResourceLinks\"",
                        "xml:base=\"res00006\"",
                        "identifier=\"res00006\"",
                        "type=\"course/x-bb-csresourcelinks\" />",
                    "<resource",
                        "bb:file=\"res00007.dat\"",
                        "bb:title=\"CourseRubricAssociation\"",
                        "xml:base=\"res00007\"",
                        "identifier=\"res00007\"",
                        "type=\"course/x-bb-crsrubricassocation\" />",
                    "</resources>",
                    "</manifest>"
                };
                string path = savePath + "\\imsmanifest.xml";
                using (StreamWriter imsmanifest = new StreamWriter(path))
                {
                    foreach (string line in manifestwithtopics)
                        imsmanifest.WriteLine(line);
                }
                fileList.Add(path, "imsmanifest.xml");
            }
            //FileStream imsmanifestpath = CreateFile(savePath, "imsmanifest.xml");
        }
        public void Createres00001()
        {
            string path = savePath + "\\res00001.dat";
            fileList.Add(path, "res00001.dat");
            using (StreamWriter res00001 = new StreamWriter(path))
            {
                //FileStream res00001 = CreateFile(savePath, "res00001.dat");
                foreach (string line in res0001assessdata)
                    res00001.WriteLine(line);

                foreach (var question in QuestionFormLoader.questionList)
                {
                    switch (question.QuestionType.InnerText)
                    {
                        case "True or False Question":
                            questionType = "True/False";
                            rcardinality = "Single";
                            break;
                        case "Multiple Answer Question":
                            questionType = "Multiple Answer";
                            rcardinality = "Multiple";
                            break;
                        case "Multiple Choice Question":
                            questionType = "Multiple Choice";
                            rcardinality = "Single";
                            break;
                        case "Essay Question":
                            questionType = "Essay";
                            rcardinality = "Single";
                            break;
                    }
                    if (Form1.AnswerNegativePointsEnabled == "true" && Form1.OverallNegativeScore == "true")
                    {
                        negativepointsind = "Y";
                    }
                    else if (Form1.AnswerNegativePointsEnabled == "true" && Form1.OverallNegativeScore == "false")
                    {
                        negativepointsind = "Q";
                    }
                    else
                    {
                        negativepointsind = "N";
                    }
                    string[] res0001itemmetadata =
                    {
                    "<item title=\"" + Form1.TestName + "_" + question.QuestionNumber + "\" maxattempts=\"0\">",
                    "<itemmetadata>",
                    "<bbmd_asi_object_id>question_" + question.QuestionNumber + "</bbmd_asi_object_id>",
                    "<bbmd_asitype>Item</bbmd_asitype>",
                    "<bbmd_assessmenttype>Test</bbmd_assessmenttype>",
                    "<bbmd_sectiontype>Subsection</bbmd_sectiontype>",
                    "<bbmd_questiontype>" + questionType + "</bbmd_questiontype>",
                    "<bbmd_is_from_cartridge>false</bbmd_is_from_cartridge>",
                    "<bbmd_is_disabled>false</bbmd_is_disabled>",
                    "<bbmd_negative_points_ind>" + negativepointsind + "</bbmd_negative_points_ind>",
                    "<bbmd_canvas_fullcrdt_ind>false</bbmd_canvas_fullcrdt_ind>",
                    "<bbmd_all_fullcredit_ind>false</bbmd_all_fullcredit_ind>",
                    "<bbmd_numbertype>letter_lower</bbmd_numbertype>",
                    "<bbmd_partialcredit>" + Form1.AnswerPartialCreditEnabled + "</bbmd_partialcredit>",
                    "<bbmd_orientationtype>vertical</bbmd_orientationtype>",
                    "<bbmd_is_extracredit>false</bbmd_is_extracredit>",
                    "<qmd_absolutescore_max>" + Form1.QuestionScore + "</qmd_absolutescore_max>",
                    "<qmd_weighting>0</qmd_weighting>",
                    "<qmd_instructornotes>",
                    "</qmd_instructornotes>",
                    "</itemmetadata>",
                    "<presentation>",
                    "<flow class=\"Block\">",
                    "<flow class=\"QUESTION_BLOCK\">",
                    "<flow class=\"FORMATTED_TEXT_BLOCK\">",
                    "<material>",
                    "<mat_extension>",
                    "<mat_formattedtext type=\"HTML\">"
                    };
                    foreach (string line in res0001itemmetadata)
                        res00001.WriteLine(line);
                    foreach (Paragraph paragraph in question.QuestionTextElements)
                    {
                        if (paragraph.InnerXml.Contains("Drawing"))
                        {
                            res00001.WriteLine("&lt;p&gt;&lt;img style=&quot;border: 0px solid rgb(0, 0, 0);&quot; alt=&quot;image00" + imagenumber + "&quot; title=&quot;image00" + imagenumber + "&quot; src=&quot;@X@EmbeddedFile.requestUrlStub@X@bbcswebdav/xid-000000" + imagenumber + "_1&quot;  /&gt;&lt;/p&gt;");
                            imagenumber += 1;
                        }
                        else
                        {
                            res00001.WriteLine("&lt;p&gt;" + paragraph.InnerText + "&lt;/p&gt;");
                        }
                    }
                    string[] endQuestionTextBlock =
                    {
                        "</mat_formattedtext>",
                        "</mat_extension>",
                        "</material>",
                        "</flow>",
                        "</flow>"
                    };
                    foreach (string line in endQuestionTextBlock)
                        res00001.WriteLine(line);
                    if (questionType == "Essay")
                    {
                        string[] responseBlockStart =
{
                            "<flow class=\"RESPONSE_BLOCK\">",
                            "<response_str ident=\"response\" rcardinality=\"" + rcardinality + "\" rtiming=\"No\">",
                            "<render_fib charset=\"us-ascii\" encoding=\"UTF_8\" rows=\"8\" columns=\"127\" maxchars=\"0\" prompt=\"Box\" fibtype=\"String\" minnumber=\"0\" maxnumber=\"0\"/>"
                        };
                        foreach (string line in responseBlockStart)
                            res00001.WriteLine(line);
                    }
                    else
                    {
                        string[] responseBlockStart =
                        {
                            "<flow class=\"RESPONSE_BLOCK\">",
                            "<response_lid ident=\"response\" rcardinality=\"" + rcardinality + "\" rtiming=\"No\">",
                            "<render_choice shuffle=\"" + Form1.AnswerRandomOrderEnabled + "\" minnumber=\"0\" maxnumber=\"0\">"
                        };
                        foreach (string line in responseBlockStart)
                            res00001.WriteLine(line);
                    }
                    if (questionType == "True/False")
                    {
                        string[] TFanswerStart =
                        {
                            "<flow_label class=\"Block\">",
                            "<response_label ident=\"true\" shuffle=\"Yes\" rarea=\"Ellipse\" rrange=\"Exact\">",
                            "<flow_mat class=\"Block\">",
                            "<material>",
                            "<mattext charset=\"us-ascii\" texttype=\"text/plain\" xml:space=\"default\">true</mattext>",
                            "</material>",
                            "</flow_mat>",
                            "</response_label>",
                            "<response_label ident=\"false\" shuffle=\"Yes\" rarea=\"Ellipse\" rrange=\"Exact\">",
                            "<flow_mat class=\"Block\">",
                            "<material>",
                            "<mattext charset=\"us-ascii\" texttype=\"text/plain\" xml:space=\"default\">false</mattext>",
                            "</material>",
                            "</flow_mat>",
                            "</response_label>",
                            "</flow_label>"
                        };
                        foreach (string line in TFanswerStart)
                            res00001.WriteLine(line);
                    }
                    if (questionType != "Essay")
                    {
                        foreach (List<Paragraph> list in question.ListOfIndividualAnswerParagraphLists)
                        {
                            string[] answerStart =
                            {
                            "<flow_label class=\"Block\">",
                            "<response_label ident=\"answer_" + (question.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1) + "\" shuffle=\"Yes\" rarea=\"Ellipse\" rrange=\"Exact\">",
                            "<flow_mat class=\"FORMATTED_TEXT_BLOCK\">",
                            "<material>",
                            "<mat_extension>",
                            "<mat_formattedtext type=\"HTML\">"
                            };
                            foreach (string line in answerStart)
                                res00001.WriteLine(line);
                            foreach (OpenXmlElement answer in list)
                            {
                                if (answer.InnerXml.Contains("Drawing"))
                                {
                                    res00001.WriteLine("&lt;p&gt;&lt;img style=&quot;border: 0px solid rgb(0, 0, 0);&quot; alt=&quot;image00" + imagenumber + "&quot; title=&quot;image00" + imagenumber + "&quot; src=&quot;@X@EmbeddedFile.requestUrlStub@X@bbcswebdav/xid-000000" + imagenumber + "_1&quot;  /&gt;&lt;/p&gt;");
                                    imagenumber += 1;
                                }
                                else
                                {
                                    res00001.WriteLine("&lt;p&gt;" + answer.InnerText + "&lt;/p&gt;");
                                }
                            }
                            string[] answerEnd =
                            {
                            "</mat_formattedtext>",
                            "</mat_extension>",
                            "</material>",
                            "</flow_mat>",
                            "</response_label>",
                            "</flow_label>"
                            };
                            foreach (string line in answerEnd)
                                res00001.WriteLine(line);
                        }
                    }
                    if (questionType == "Essay")
                    {
                        string[] responseBlockEnd =
                        {
                            "</response_str>",
                            "</flow>",
                            "</flow>",
                            "</presentation>"
                        };
                        foreach (string line in responseBlockEnd)
                            res00001.WriteLine(line);
                    }
                    else
                    {
                        string[] responseBlockEnd =
                        {
                            "</render_choice>",
                            "</response_lid>",
                            "</flow>",
                            "</flow>",
                            "</presentation>"
                        };
                        foreach (string line in responseBlockEnd)
                            res00001.WriteLine(line);
                    }
                    if (questionType == "Multiple Answer")
                    {
                        string[] questionEvaluationStart =
                        {
                        "<resprocessing scoremodel=\"SumOfScores\">",
                        "<outcomes>",
                        "<decvar varname=\"SCORE\" vartype=\"Decimal\" defaultval=\"0.0\" minvalue=\"0.0\" maxvalue=\"" + Form1.QuestionScore + "\" />",
                        "</outcomes>",
                        "<respcondition title=\"correct\">",
                        "<conditionvar>",
                        "<and>"
                        };
                        foreach (string line in questionEvaluationStart)
                            res00001.WriteLine(line);
                        foreach (List<Paragraph> list in question.ListOfIndividualAnswerParagraphLists)
                        {
                            if (list.Any(or => or.Descendants<Color>().Any()))
                            {
                                List<string> respident = new List<string>();
                                respident.Add("<varequal respident=\"response\" case=\"No\">answer_" + (question.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1) + "</varequal>");
                                foreach (string line in respident)
                                    res00001.WriteLine(line);
                            }
                            else
                            {
                                List<string> respident = new List<string>();
                                respident.Add("<not>");
                                respident.Add("<varequal respident=\"response\" case=\"No\">answer_" + (question.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1) + "</varequal>");
                                respident.Add("</not>");
                                foreach (string line in respident)
                                    res00001.WriteLine(line);
                            }

                        }
                        string[] questionEvaluationMid =
                        {
                            "</and>",
                            "</conditionvar>",
                            "<setvar variablename=\"SCORE\" action=\"Set\">SCORE.max</setvar>",
                            "<displayfeedback linkrefid=\"correct\" feedbacktype=\"Response\" />",
                            "</respcondition>",
                            "<respcondition title=\"incorrect\">",
                            "<conditionvar>",
                            "<other />",
                            "</conditionvar>",
                            "<setvar variablename=\"SCORE\" action=\"Set\">0.0</setvar>",
                            "<displayfeedback linkrefid=\"incorrect\" feedbacktype=\"Response\" />",
                            "</respcondition>"
                        };
                        foreach (string line in questionEvaluationMid)
                            res00001.WriteLine(line);
                        foreach (List<Paragraph> list in question.ListOfIndividualAnswerParagraphLists)
                        {
                            res00001.WriteLine("<respcondition>");
                            res00001.WriteLine("<conditionvar>");
                            res00001.WriteLine("<varequal respident=\"answer_" + (question.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1) + "\" case=\"No\" />");
                            res00001.WriteLine("</conditionvar>");
                            if (Form1.AnswerNegativePointsEnabled == "true")
                            {
                                if (list.Any(or => or.Descendants<Color>().Any()))
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">" + 100 / question.CorrectAnswers.Count() + "</setvar>");
                                }
                                else
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">-" + 100 / (question.ListOfIndividualAnswerParagraphLists.Count() - question.CorrectAnswers.Count()) + "</setvar>");
                                }
                            }
                            if (Form1.AnswerPartialCreditEnabled == "true")
                            {
                                if (list.Any(or => or.Descendants<Color>().Any()))
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">" + 100 / question.CorrectAnswers.Count() + "</setvar>");
                                }
                                else
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">0</setvar>");
                                }
                            }
                            if (Form1.AnswerNegativePointsEnabled == "true" && Form1.AnswerPartialCreditEnabled == "true")
                            {
                                if (list.Any(or => or.Descendants<Color>().Any()))
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">" + 100 / question.CorrectAnswers.Count() + "</setvar>");
                                }
                                else
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">-" + 100 / (question.ListOfIndividualAnswerParagraphLists.Count() - question.CorrectAnswers.Count()) + "</setvar>");
                                }
                            }
                            else if (Form1.AnswerNegativePointsEnabled == "false" && Form1.AnswerPartialCreditEnabled == "false")
                            {
                                res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">0</setvar>");
                            }
                            res00001.WriteLine("<displayfeedback linkrefid=\"answer_" + (question.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1) + "\" feedbacktype=\"Response\"/>");
                            res00001.WriteLine("</respcondition>");
                        }
                        res00001.WriteLine("</resprocessing>");
                    }
                    if (questionType == "Multiple Choice")
                    {
                        string[] questionEvaluationStart =
                        {
                        "<resprocessing scoremodel=\"SumOfScores\">",
                        "<outcomes>",
                        "<decvar varname=\"SCORE\" vartype=\"Decimal\" defaultval=\"0.0\" minvalue=\"0.0\" maxvalue=\"" + Form1.QuestionScore + "\" />",
                        "</outcomes>",
                        "<respcondition title=\"correct\">",
                        "<conditionvar>"
                        };
                        foreach (string line in questionEvaluationStart)
                            res00001.WriteLine(line);
                        foreach (List<Paragraph> list in question.ListOfIndividualAnswerParagraphLists)
                        {
                            foreach (OpenXmlElement answer in list)
                            {
                                if (answer.Descendants<Color>().Any())
                                {
                                    correctAnswer = "answer_" + (question.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1);
                                }
                            }
                        }
                        res00001.WriteLine("<varequal respident=\"response\" case=\"No\">" + correctAnswer + "</varequal>");
                        string[] questionEvaluationMid =
                        {
                            "</conditionvar>",
                            "<setvar variablename=\"SCORE\" action=\"Set\">SCORE.max</setvar>",
                            "<displayfeedback linkrefid=\"correct\" feedbacktype=\"Response\" />",
                            "</respcondition>",
                            "<respcondition title=\"incorrect\">",
                            "<conditionvar>",
                            "<other />",
                            "</conditionvar>",
                            "<setvar variablename=\"SCORE\" action=\"Set\">0.0</setvar>",
                            "<displayfeedback linkrefid=\"incorrect\" feedbacktype=\"Response\" />",
                            "</respcondition>"
                        };
                        foreach (string line in questionEvaluationMid)
                            res00001.WriteLine(line);
                        foreach (List<Paragraph> list in question.ListOfIndividualAnswerParagraphLists)
                        {
                            res00001.WriteLine("<respcondition>");
                            res00001.WriteLine("<conditionvar>");
                            res00001.WriteLine("<varequal respident=\"answer_" + (question.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1) + "\" case=\"No\" />");
                            res00001.WriteLine("</conditionvar>");
                            if (Form1.AnswerNegativePointsEnabled == "true")
                            {
                                if (list.Any(or => or.Descendants<Color>().Any()))
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">100</setvar>");
                                }
                                else
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">-" + 100 / (question.ListOfIndividualAnswerParagraphLists.Count() - question.CorrectAnswers.Count()) + "</setvar>");
                                }
                            }
                            else if (Form1.AnswerNegativePointsEnabled == "false")
                            {
                                if (list.Any(or => or.Descendants<Color>().Any()))
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">100</setvar>");
                                }
                                else
                                {
                                    res00001.WriteLine("<setvar variablename=\"SCORE\" action=\"Set\">0</setvar>");
                                }
                            }
                            res00001.WriteLine("<displayfeedback linkrefid=\"answer_" + (question.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1) + "\" feedbacktype=\"Response\" />");
                            res00001.WriteLine("</respcondition>");
                        }
                        res00001.WriteLine("</resprocessing>");
                    }
                    if (questionType == "True/False")
                    {
                        string[] TFResponseBlock =
                        {
                            "<resprocessing scoremodel=\"SumOfScores\">",
                            "<outcomes>",
                            "<decvar varname=\"SCORE\" vartype=\"Decimal\" defaultval=\"0\" minvalue=\"0\" maxvalue=\"" + Form1.QuestionScore + "\"/>",
                            "</outcomes>",
                            "<respcondition title=\"correct\">",
                            "<conditionvar>",
                            "<varequal respident=\"response\" case=\"No\">" + question.IndividualAnswerParagraphs[0].InnerText + "</varequal>",
                            "</conditionvar>",
                            "<setvar variablename=\"SCORE\" action=\"Set\">SCORE.max</setvar>",
                            "<displayfeedback linkrefid=\"correct\" feedbacktype=\"Response\"/>",
                            "</respcondition>",
                            "<respcondition title=\"incorrect\">",
                            "<conditionvar>",
                            "<other/>",
                            "</conditionvar>",
                            "<setvar variablename=\"SCORE\" action=\"Set\">0</setvar>",
                            "<displayfeedback linkrefid=\"incorrect\" feedbacktype=\"Response\"/>",
                            "</respcondition>",
                            "</resprocessing>"
                        };
                        foreach (string line in TFResponseBlock)
                            res00001.WriteLine(line);
                    }
                    if (questionType == "Essay")
                    {
                        string[] responseBlock =
                        {
                            "<resprocessing scoremodel=\"SumOfScores\">",
                            "<outcomes>",
                            "<decvar varname=\"SCORE\" vartype=\"Decimal\" defaultval=\"0\" minvalue=\"0\" maxvalue=\"" + Form1.QuestionScore + "\" />",
                            "</outcomes>",
                            "<respcondition title=\"correct\">",
                            "<conditionvar/>",
                            "<setvar variablename=\"SCORE\" action=\"Set\">" + Form1.QuestionScore + "</setvar>",
                            "<displayfeedback linkrefid=\"correct\" feedbacktype=\"Response\"/>",
                            "</respcondition>",
                            "<respcondition title=\"incorrect\">",
                            "<conditionvar>",
                            "<other/>",
                            "</conditionvar>",
                            "<setvar variablename=\"SCORE\" action=\"Set\">0</setvar>",
                            "<displayfeedback linkrefid=\"incorrect\" feedbacktype=\"Response\"/>",
                            "</respcondition>",
                            "</resprocessing>"
                        };
                        foreach (string line in responseBlock)
                            res00001.WriteLine(line);
                    };
                    string[] itemFeedback =
                    {
                        "<itemfeedback ident=\"correct\" view=\"All\">",
                        "<flow_mat class=\"Block\">",
                        "<flow_mat class=\"FORMATTED_TEXT_BLOCK\">",
                        "<material>",
                        "<mat_extension>",
                        "<mat_formattedtext type=\"HTML\">Correct</mat_formattedtext>",
                        "</mat_extension>",
                        "</material>",
                        "</flow_mat>",
                        "</flow_mat>",
                        "</itemfeedback>",
                        "<itemfeedback ident=\"incorrect\" view=\"All\">",
                        "<flow_mat class=\"Block\">",
                        "<flow_mat class=\"FORMATTED_TEXT_BLOCK\">",
                        "<material>",
                        "<mat_extension>",
                        "<mat_formattedtext type=\"HTML\">Incorrect</mat_formattedtext>",
                        "</mat_extension>",
                        "</material>",
                        "</flow_mat>",
                        "</flow_mat>",
                        "</itemfeedback>"
                    };
                    foreach (string line in itemFeedback)
                        res00001.WriteLine(line);
                    if (questionType == "Multiple Choice" || questionType == "Multiple Answer")
                    {
                        foreach (List<Paragraph> list in question.ListOfIndividualAnswerParagraphLists)
                        {
                            string[] individualAnswerFeedbackpt1 =
                            {
                                "<itemfeedback ident=\"answer_" + (question.ListOfIndividualAnswerParagraphLists.IndexOf(list) + 1) + "\" view=\"All\">",
                                "<solution view=\"All\" feedbackstyle=\"Complete\">",
                                "<solutionmaterial>",
                                "<flow_mat class=\"Block\">",
                                "<flow_mat class=\"FORMATTED_TEXT_BLOCK\">",
                                "<material>",
                                "<mat_extension>"
                            };
                            foreach (string line in individualAnswerFeedbackpt1)
                                res00001.WriteLine(line);
                            foreach (OpenXmlElement answer in list)
                            {
                                if (answer.Descendants<Color>().Any())
                                {
                                    answerResult = "correct";
                                }
                                else
                                {
                                    answerResult = "incorrect";
                                }
                            }
                            string[] individualAnswerFeedbackpt2 =
                            {
                                "<mat_formattedtext type=\"HTML\">" + answerResult + "</mat_formattedtext>",
                                "</mat_extension>",
                                "</material>",
                                "</flow_mat>",
                                "</flow_mat>",
                                "</solutionmaterial>",
                                "</solution>",
                                "</itemfeedback>"
                            };
                            foreach (string line in individualAnswerFeedbackpt2)
                                res00001.WriteLine(line);
                        }
                    }
                    if (questionType == "Essay")
                    {
                        string[] answerFeedback =
                        {
                            "<itemfeedback ident=\"solution\" view=\"All\">",
                            "<solution view=\"All\" feedbackstyle=\"Complete\">",
                            "<solutionmaterial>",
                            "<flow_mat class=\"Block\">",
                            "<material>",
                            "<mat_extension>",
                            "<mat_formattedtext type=\"HTML\"/>",
                            "</mat_extension>",
                            "</material>",
                            "</flow_mat>",
                            "</solutionmaterial>",
                            "</solution>",
                            "</itemfeedback>"
                        };
                        foreach (string line in answerFeedback)
                            res00001.WriteLine(line);
                    }
                    res00001.WriteLine("</item>");
                }
                foreach(string line in res0001assessdataend)
                    res00001.WriteLine(line);
            }
        }
        public void Createres00002()
        {
            if(QuestionFormLoader.QuestionTopics.Count() != 0 || QuestionFormLoader.QuestionDifficulty.Count() != 0)
            { 
                string path = savePath + "\\res00002.dat";
                fileList.Add(path, "res00002.dat");
                using (StreamWriter res00002 = new StreamWriter(path))
                {
                    res00002.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                    res00002.WriteLine("<CATEGORIES>");
                    foreach(Text value in QuestionFormLoader.QuestionTopics)
                    { 
                        string[] lines =
                        {
                            "<CATEGORY id=\"" + value.InnerText + "\">",
                            "<TITLE>" + value.InnerText + "</TITLE>",
                            "<TYPE>learning_objective</TYPE>",
                            "<COURSEID",
                            "value =\"FAKE-COURSE\"/>",
                            "</CATEGORY>"
                        };
                        foreach (string line in lines)
                            res00002.WriteLine(line);
                    }
                    foreach (Text value in QuestionFormLoader.QuestionDifficulty)
                    {
                        string[] lines =
                        {
                            "<CATEGORY id=\"" + value.InnerText + "\">",
                            "<TITLE>" + value.InnerText + "</TITLE>",
                            "<TYPE>level_of_difficulty</TYPE>",
                            "<COURSEID",
                            "value =\"FAKE-COURSE\"/>",
                            "</CATEGORY>"
                        };
                        foreach (string line in lines)
                            res00002.WriteLine(line);
                    }
                    res00002.WriteLine("</CATEGORIES>");
                }
            }
            else
            {
                return;
            }
        }
        public void Createres00003()
        {
            if (QuestionFormLoader.QuestionTopics.Count() != 0 || QuestionFormLoader.QuestionDifficulty.Count() != 0)
            { 
                string path = savePath + "\\res00003.dat";
                fileList.Add(path, "res00003.dat");
                using (StreamWriter res00003 = new StreamWriter(path))
                {
                    int i = 1;
                    res00003.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                    res00003.WriteLine("<ITEMCATEGORIES>");
                    foreach(var question in QuestionFormLoader.questionList)
                    { 
                        foreach (var topic in question.Topics)
                        {
                            
                            string[] lines =
                            {
                                "<ITEMCATEGORY id=\"_000000" + i +"_1\">",
                                "<CATEGORYID value=\"" + topic.Key.InnerText + "\"/>",
                                "<QUESTIONID value=\"question_" + topic.Value + "\"/>",
                                "</ITEMCATEGORY>"
                            };
                            foreach (string line in lines)
                                res00003.WriteLine(line);
                            i++;
                        }
                        foreach (var topic in question.Difficulty)
                        {

                            string[] lines =
                            {
                                "<ITEMCATEGORY id=\"_000000" + i +"_1\">",
                                "<CATEGORYID value=\"" + topic.Key.InnerText + "\"/>",
                                "<QUESTIONID value=\"question_" + topic.Value + "\"/>",
                                "</ITEMCATEGORY>"
                            };
                            foreach (string line in lines)
                                res00003.WriteLine(line);
                            i++;
                        }
                    }
                    res00003.WriteLine("</ITEMCATEGORIES>");
                }
            }
            else
            {
                return;
            }
        }
        public void Createres00004()
        {
            string[] lines =
            {
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<ASSESSMENTCREATIONSETTINGS>",
                "<ASSESSMENTCREATIONSETTING id=\"_0000_1\">",
                "<QTIASSESSMENTID value=\"" + Form1.TestName + " " + DateTime.Today + "\"/>",
                "<ANSWERFEEDBACKENABLED>true</ANSWERFEEDBACKENABLED>",
                "<QUESTIONATTACHMENTSENABLED>true</QUESTIONATTACHMENTSENABLED>",
                "<ANSWERATTACHMENTSENABLED>true</ANSWERATTACHMENTSENABLED>",
                "<QUESTIONMETADATAENABLED>true</QUESTIONMETADATAENABLED>",
                "<DEFAULTPOINTVALUEENABLED>" + Form1.DefaultScore + "</DEFAULTPOINTVALUEENABLED>",
                "<DEFAULTPOINTVALUE>" + Form1.QuestionScore + "</DEFAULTPOINTVALUE>",
                "<ANSWERPARTIALCREDITENABLED>true</ANSWERPARTIALCREDITENABLED>",
                "<ANSWERNEGATIVEPOINTSENABLED>" + Form1.AnswerNegativePointsEnabled + "</ANSWERNEGATIVEPOINTSENABLED>",
                "<ANSWERRANDOMORDERENABLED>" + Form1.AnswerRandomOrderEnabled + "</ANSWERRANDOMORDERENABLED>",
                "<ANSWERORIENTATIONENABLED>true</ANSWERORIENTATIONENABLED>",
                "<ANSWERNUMBEROPTIONSENABLED>true</ANSWERNUMBEROPTIONSENABLED>",
                "<USEPOINTSFROMSOURCEBYDEFAULT>true</USEPOINTSFROMSOURCEBYDEFAULT>",
                "</ASSESSMENTCREATIONSETTING>",
                "</ASSESSMENTCREATIONSETTINGS>"
            };
            //FileStream res00002 = CreateFile(savePath, "res00002.dat");
            string path = string.Empty;
            if (QuestionFormLoader.QuestionTopics.Count() == 0 && QuestionFormLoader.QuestionDifficulty.Count() == 0)
            { 
                path = savePath + "\\res00002.dat";
                fileList.Add(path, "res00002.dat");
            }
            else
            {
                path = savePath + "\\res00004.dat";
                fileList.Add(path, "res00004.dat");
            }
            using (StreamWriter res00004 = new StreamWriter(path))
            {
                foreach (string line in lines)
                    res00004.WriteLine(line);
            }
        }
        public void Createres00005to7()
        {
            String[] res00005lines =
            {
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<LEARNRUBRICS/>"
            };
            String[] res00006lines =
                {
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<cms_resource_link_list/>"
            };
            String[] res00007lines =
            {
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<COURSERUBRICASSOCIATIONS/>"
            };
            string res00005path = string.Empty;
            if (QuestionFormLoader.QuestionTopics.Count() == 0 && QuestionFormLoader.QuestionDifficulty.Count() == 0)
            { 
                res00005path = savePath + "\\res00003.dat";
                fileList.Add(res00005path, "res00003.dat");
            }
            else
            {
                res00005path = savePath + "\\res00005.dat";
                fileList.Add(res00005path, "res00005.dat");
            }
            using (StreamWriter res00005 = new StreamWriter(res00005path))
            {
                foreach (string line in res00005lines)
                    res00005.WriteLine(line);
            }
            string res00006path = string.Empty;
            if (QuestionFormLoader.QuestionTopics.Count() == 0 && QuestionFormLoader.QuestionDifficulty.Count() == 0)
            {
                res00006path = savePath + "\\res00004.dat";
                fileList.Add(res00006path, "res00004.dat");
            }
            else
            {
                res00006path = savePath + "\\res00006.dat";
                fileList.Add(res00006path, "res00006.dat");
            }
            using (StreamWriter res00006 = new StreamWriter(res00006path))
            {
                foreach (string line in res00006lines)
                    res00006.WriteLine(line);
            }
            string res00007path = string.Empty;
            if (QuestionFormLoader.QuestionTopics.Count() == 0 && QuestionFormLoader.QuestionDifficulty.Count() == 0)
            {
                res00007path = savePath + "\\res00005.dat";
                fileList.Add(res00007path, "res00005.dat");
            }
            else
            {
                res00007path = savePath + "\\res00007.dat";
                fileList.Add(res00007path, "res00007.dat");
            }
            using (StreamWriter res00007 = new StreamWriter(res00007path))
            {
                foreach (string line in res00007lines)
                    res00007.WriteLine(line);
            }
        }
        public void Createzip()
        {
            using (FileStream testzip = new FileStream(savePath + "\\" + Form1.TestName + ".zip", FileMode.OpenOrCreate))
            {
                using (ZipArchive archive = new ZipArchive(testzip, ZipArchiveMode.Update))
                { 
                    foreach(KeyValuePair<string, string> file in fileList)
                    {
                        archive.CreateEntryFromFile(file.Key, file.Value);
                        //archive.ExtractToDirectory(savePath);
                    }
                    DirectoryInfo Images = new DirectoryInfo(savePath + "\\csfiles\\home_dir");
                    FileInfo[] files = Images.GetFiles("*");
                    foreach(FileInfo file in files)
                    {
                        archive.CreateEntryFromFile(savePath + "\\csfiles\\home_dir" + "\\" + file.Name, "csfiles\\home_dir" + "/" + file.Name);
                    }
                }
            }
        }
    }
}
