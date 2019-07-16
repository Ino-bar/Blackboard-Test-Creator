using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Blackboard_Test_Creator
{
    class TestCreator
    {
        public string savePath = Form1.TestFilePath;
        static int totalScore = QuestionFormLoader.questionList.Count() * Form1.QuestionScore;
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
            "<mat_formattedtext type = \"HTML\"/>",
            "</mat_extension>",
            "</material>",
            "</flow_mat>",
            "</rubric>",
            "<presentation_material>",
            "<flow_mat class=\"Block\">",
            "<material>",
            "<mat_extension>",
            "<mat_formattedtext type = \"HTML\"/>",
            "</mat_extension>",
            "</material>",
            "</flow_mat>",
            "</presentation_material>",
            "<section>",
            "<sectionmetadata>",
            "<bbmd_asi_object_id> section_0 </bbmd_asi_object_id>",
            "<bbmd_asitype> Section </bbmd_asitype>",
            "<bbmd_assessmenttype> Test </bbmd_assessmenttype>",
            "<bbmd_sectiontype> Subsection </bbmd_sectiontype>",
            "<bbmd_questiontype > Multiple Choice</bbmd_questiontype>",
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
        string[] res0001questionblock =
        {

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
        public void CreateBBPackage()
        {
            string path = savePath + "\\.bb-package-info";
            using (StreamWriter BBPackage = new StreamWriter(path))
            {
                BBPackage.WriteLine("cx.package.info.version=6.0");
            }
        }
        public void Createimsmanifest()
        {
            string[] lines =
            {
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<manifest identifier=\"man00001\"",
                "xmlns:bb = \"http://www.blackboard.com/content-packaging/\">",
                "<organizations default = \"toc00001\">",
                "<organization identifier = \"toc00001\"/>",
                "</organizations >",
                "<resources >",
                "<resource",
                    "bb: file = \"res00001.dat\"",
                    "bb: title = \"" + Form1.TestName + "\"",
                    "xml: base = \"res00001\"",
                    "identifier = \"res00001\"",
                    "type = \"assessment/x-bb-qti-test\"/>",
                "<resource",
                    "bb: file = \"res00002.dat\"",
                    "bb: title = \"Assessment Creation Settings\"",
                    "xml: base = \"res00002\"",
                    "identifier = \"res00002\"",
                    "type = \"course/x-bb-courseassessmentcreationsettings\"/>",
                "<resource",
                    "bb: file = \"res00003.dat\"",
                    "bb: title = \"LearnRubrics\"",
                    "xml: base = \"res00003\"",
                    "identifier = \"res00003\"",
                    "type = \"course/x-bb-rubrics\"/>",
                "<resource",
                    "bb: file = \"res00004.dat\"",
                    "bb: title = \"CSResourceLinks\"",
                    "xml: base = \"res00004\"",
                    "identifier = \"res00004\"",
                    "type = \"course/x-bb-csresourcelinks\"/>",
                "<resource",
                    "bb: file = \"res00005.dat\"",
                    "bb: title = \"CourseRubricAssociation\"",
                    "xml: base = \"res00005\"",
                    "identifier = \"res00005\"",
                    "type = \"course/x-bb-crsrubricassocation\"/>",
                "</resources>",
                "</manifest> "
            };
            //FileStream imsmanifestpath = CreateFile(savePath, "imsmanifest.xml");
            string path = savePath + "\\imsmanifest.xml";
            using (StreamWriter imsmanifest = new StreamWriter(path))
            {
                foreach (string line in lines)
                    imsmanifest.WriteLine(line);
            }
        }
        public void Createres00001()
        {
            FileStream res00001 = CreateFile(savePath, "res00001.dat");
            foreach(var question in QuestionFormLoader.questionList)
            {
                
            }
        }
        public void Createres00002()
        {
            String[] lines =
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
            string path = savePath + "\\res00002.dat";
            using (StreamWriter res00002 = new StreamWriter(path))
            {
                foreach (string line in lines)
                    res00002.WriteLine(line);
            }
        }
        public void Createres00003to5()
        {
            String[] res00003lines =
            {
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<LEARNRUBRICS/>"
            };
            String[] res00004lines =
                {
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<cms_resource_link_list/>"
            };
            String[] res00005lines =
    {
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
                "<COURSERUBRICASSOCIATIONS/>"
            };
            string res00003path = savePath + "\\res00003.dat";
            using (StreamWriter res00003 = new StreamWriter(res00003path))
            {
                foreach (string line in res00003lines)
                    res00003.WriteLine(line);
            }
            string res00004path = savePath + "\\res00004.dat";
            using (StreamWriter res00004 = new StreamWriter(res00004path))
            {
                foreach (string line in res00004lines)
                    res00004.WriteLine(line);
            }
            string res00005path = savePath + "\\res00005.dat";
            using (StreamWriter res00005 = new StreamWriter(res00005path))
            {
                foreach (string line in res00005lines)
                    res00005.WriteLine(line);
            }
        }
    }
}
