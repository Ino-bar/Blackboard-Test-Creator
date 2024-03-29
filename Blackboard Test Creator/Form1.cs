﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Tools.Word;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml;
using System.IO;

namespace Blackboard_Test_Creator
{
    public partial class Form1 : Form
    {
        public static string TestFormFilePath;
        public static string TestFilePath;
        public static string TestName;
        public static int QuestionScore;
        public static string DefaultScore;
        public static string AnswerNegativePointsEnabled;
        public static string OverallNegativeScore;
        public static string AnswerRandomOrderEnabled;
        public static string AnswerPartialCreditEnabled;
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
                    bool Locked = false;
                    try
                    {
                        System.IO.FileStream fs =
                            File.Open(TestFormFilePath, FileMode.Open,
                            FileAccess.ReadWrite, FileShare.None);
                        fs.Close();
                    }
                    catch (IOException ex)
                    {
                        Locked = true;
                    }
                    if (Locked == true)
                    {
                        MessageBox.Show("The file is currently open in another application. Please close the file and then run the program");
                        return;
                    }
                    else
                    {
                        chosenFormFilenameLabel.Text = getTestForm.FileName;
                    }
                }
                QuestionFormLoader questionFormLoader = new QuestionFormLoader();
                questionFormLoader.FormLoader();
            }
            if(!string.IsNullOrEmpty(TestFilePath) && !string.IsNullOrEmpty(TestName))
            {
                startButton.Enabled = true;
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
            if (!string.IsNullOrEmpty(TestFormFilePath) && !string.IsNullOrEmpty(TestName))
            {
                startButton.Enabled = true;
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
        void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            Thread.Sleep(2000);
        }
        void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Value = progressBar1.Minimum;
            completionIndicatorLabel.Visible = true;
        }
        private void startButton_Click(object sender, EventArgs e)
        {
            TestName = TestNameTextBox.Text;
            if (enableNegativeMarkingCheckBox.Checked == true)
            {
                AnswerNegativePointsEnabled = "true";
            }
            else if (enableNegativeMarkingCheckBox.Checked == false)
            {
                AnswerNegativePointsEnabled = "false";
            }
            if (allowPartialCreditCheckBox.Checked == true || enableNegativeMarkingCheckBox.Checked == true)
            {
                AnswerPartialCreditEnabled = "true";
            }
            else if (allowPartialCreditCheckBox.Checked == false || enableNegativeMarkingCheckBox.Checked == false)
            {
                AnswerPartialCreditEnabled = "false";
            }
            if(allowOverallNegativeScoreCheckBox.Checked == true)
            {
                OverallNegativeScore = "true";
            }
            else
            {
                OverallNegativeScore = "false";
            }
            if (!string.IsNullOrEmpty(questionScoreTextBox.Text))
            {
                QuestionScore = Int32.Parse(questionScoreTextBox.Text);
                DefaultScore = "false";
            }
            else if (string.IsNullOrEmpty(questionScoreTextBox.Text))
            {
                QuestionScore = 1;
                DefaultScore = "true";
            }
            if (AnswerRandomOrderEnabledCheckBox.Checked == true)
            {
                AnswerRandomOrderEnabled = "Yes";
            }
            else if (AnswerRandomOrderEnabledCheckBox.Checked == false)
            {
                AnswerRandomOrderEnabled = "No";
            }
            TestCreator testCreator = new TestCreator();
            progressBar1.MarqueeAnimationSpeed = 50;
            BackgroundWorker bw = new BackgroundWorker();
            bw.DoWork += bw_DoWork;
            bw.RunWorkerCompleted += bw_RunWorkerCompleted;
            bw.RunWorkerAsync();
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

        private void TestNameTextBox_TextChanged(object sender, EventArgs e)
        {
            TestName = TestNameTextBox.Text;
            if (!string.IsNullOrEmpty(TestFormFilePath) && !string.IsNullOrEmpty(TestFilePath))
            {
                startButton.Enabled = true;
            }
        }
    }
}
