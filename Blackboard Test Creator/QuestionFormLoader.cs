using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.ComponentModel;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace Blackboard_Test_Creator
{
    class QuestionFormLoader
    {
        public string formPath = Form1.TestFormFilePath;
        public static Word.Application wordApp = new Word.Application();
        public static Word.Document Form;
        public void FormLoader()
        {
            if (formPath != null)
            { 
                Form = wordApp.Documents.Open(formPath);
            }
        }
    }
}
