using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Blackboard_Test_Creator
{
    public partial class Form1 : Form
    {
        public static string TestFormFilePath;
        public static string TestFilePath;
        public Form1()
        {
            InitializeComponent();
        }

        private void chooseFormButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog getTestForm= new OpenFileDialog())
            {
                getTestForm.Title = "Please choose the student data table";
                getTestForm.InitialDirectory = "c:\\";
                getTestForm.Filter = "Word files(*.docx)|*.docx";
                getTestForm.RestoreDirectory = true;
                if (getTestForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TestFormFilePath = getTestForm.FileName;
                    chosenFormFilenameLabel.Text = getTestForm.FileName;
                }
                QuestionFormLoader questionFormLoader = new QuestionFormLoader();
                questionFormLoader.FormLoader();
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            if (QuestionFormLoader.wordprocessingDocument != null)
            {
                //QuestionFormLoader.wordprocessingDocument.Close();
                //QuestionFormLoader.stream.Close();
                //QuestionFormLoader.wordprocessingDocument.Dispose();
            }
            this.Close();
        }

        public void chooseSavePathButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog SetTestSaveLocation = new FolderBrowserDialog())
            {
                SetTestSaveLocation.RootFolder = Environment.SpecialFolder.Desktop;
                if (SetTestSaveLocation.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TestFilePath = SetTestSaveLocation.SelectedPath;
                    savePathTextBox.Text = TestFilePath;
                }
            }
        }

        private void questionScoreTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            questionScoreTextBox.MaxLength = 2;
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            TestCreator testCreator = new TestCreator();
            var methods = testCreator.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly | BindingFlags.Static);
            for(int i = 1; i < methods.Count(); i++)
            {
                methods[i].Invoke(testCreator, null);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (QuestionFormLoader.wordprocessingDocument != null)
            {
                //QuestionFormLoader.wordprocessingDocument.Close();
                //QuestionFormLoader.stream.Close();
                //QuestionFormLoader.wordprocessingDocument.Dispose();
            }
        }
    }
}
