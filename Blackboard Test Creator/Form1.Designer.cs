namespace Blackboard_Test_Creator
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.chooseFileLabel = new System.Windows.Forms.Label();
            this.chooseFormButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.startButton = new System.Windows.Forms.Button();
            this.chooseSavePathButton = new System.Windows.Forms.Button();
            this.savePathTextBox = new System.Windows.Forms.TextBox();
            this.testSettingsPanel = new System.Windows.Forms.Panel();
            this.questionScoreLabel = new System.Windows.Forms.Label();
            this.questionScoreTextBox = new System.Windows.Forms.TextBox();
            this.allowOverallNegativeScoreCheckBox = new System.Windows.Forms.CheckBox();
            this.enableNegativeMarkingCheckBox = new System.Windows.Forms.CheckBox();
            this.chosenFormFilenameLabel = new System.Windows.Forms.Label();
            this.testSettingsPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // chooseFileLabel
            // 
            this.chooseFileLabel.AutoSize = true;
            this.chooseFileLabel.Location = new System.Drawing.Point(13, 13);
            this.chooseFileLabel.Name = "chooseFileLabel";
            this.chooseFileLabel.Size = new System.Drawing.Size(121, 13);
            this.chooseFileLabel.TabIndex = 0;
            this.chooseFileLabel.Text = "Choose a question form:";
            // 
            // chooseFormButton
            // 
            this.chooseFormButton.Location = new System.Drawing.Point(13, 30);
            this.chooseFormButton.Name = "chooseFormButton";
            this.chooseFormButton.Size = new System.Drawing.Size(121, 23);
            this.chooseFormButton.TabIndex = 1;
            this.chooseFormButton.Text = "Choose";
            this.chooseFormButton.UseVisualStyleBackColor = true;
            this.chooseFormButton.Click += new System.EventHandler(this.chooseFormButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(390, 285);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // startButton
            // 
            this.startButton.Location = new System.Drawing.Point(309, 285);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(75, 23);
            this.startButton.TabIndex = 3;
            this.startButton.Text = "Start";
            this.startButton.UseVisualStyleBackColor = true;
            this.startButton.Click += new System.EventHandler(this.startButton_Click);
            // 
            // chooseSavePathButton
            // 
            this.chooseSavePathButton.Location = new System.Drawing.Point(13, 285);
            this.chooseSavePathButton.Name = "chooseSavePathButton";
            this.chooseSavePathButton.Size = new System.Drawing.Size(75, 23);
            this.chooseSavePathButton.TabIndex = 4;
            this.chooseSavePathButton.Text = "Choose";
            this.chooseSavePathButton.UseVisualStyleBackColor = true;
            this.chooseSavePathButton.Click += new System.EventHandler(this.chooseSavePathButton_Click);
            // 
            // savePathTextBox
            // 
            this.savePathTextBox.Location = new System.Drawing.Point(94, 287);
            this.savePathTextBox.Name = "savePathTextBox";
            this.savePathTextBox.ReadOnly = true;
            this.savePathTextBox.Size = new System.Drawing.Size(199, 20);
            this.savePathTextBox.TabIndex = 5;
            // 
            // testSettingsPanel
            // 
            this.testSettingsPanel.Controls.Add(this.questionScoreLabel);
            this.testSettingsPanel.Controls.Add(this.questionScoreTextBox);
            this.testSettingsPanel.Controls.Add(this.allowOverallNegativeScoreCheckBox);
            this.testSettingsPanel.Controls.Add(this.enableNegativeMarkingCheckBox);
            this.testSettingsPanel.Location = new System.Drawing.Point(16, 73);
            this.testSettingsPanel.Name = "testSettingsPanel";
            this.testSettingsPanel.Size = new System.Drawing.Size(231, 173);
            this.testSettingsPanel.TabIndex = 6;
            // 
            // questionScoreLabel
            // 
            this.questionScoreLabel.AutoSize = true;
            this.questionScoreLabel.Cursor = System.Windows.Forms.Cursors.SizeAll;
            this.questionScoreLabel.Location = new System.Drawing.Point(43, 64);
            this.questionScoreLabel.Name = "questionScoreLabel";
            this.questionScoreLabel.Size = new System.Drawing.Size(182, 13);
            this.questionScoreLabel.TabIndex = 3;
            this.questionScoreLabel.Text = "Score For Each Question (default 10)";
            // 
            // questionScoreTextBox
            // 
            this.questionScoreTextBox.Location = new System.Drawing.Point(12, 61);
            this.questionScoreTextBox.Name = "questionScoreTextBox";
            this.questionScoreTextBox.Size = new System.Drawing.Size(25, 20);
            this.questionScoreTextBox.TabIndex = 2;
            this.questionScoreTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.questionScoreTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.questionScoreTextBox_KeyPress);
            // 
            // allowOverallNegativeScoreCheckBox
            // 
            this.allowOverallNegativeScoreCheckBox.AutoSize = true;
            this.allowOverallNegativeScoreCheckBox.Location = new System.Drawing.Point(12, 37);
            this.allowOverallNegativeScoreCheckBox.Name = "allowOverallNegativeScoreCheckBox";
            this.allowOverallNegativeScoreCheckBox.Size = new System.Drawing.Size(164, 17);
            this.allowOverallNegativeScoreCheckBox.TabIndex = 1;
            this.allowOverallNegativeScoreCheckBox.Text = "Allow Overall Negative Score";
            this.allowOverallNegativeScoreCheckBox.UseVisualStyleBackColor = true;
            // 
            // enableNegativeMarkingCheckBox
            // 
            this.enableNegativeMarkingCheckBox.AutoSize = true;
            this.enableNegativeMarkingCheckBox.Location = new System.Drawing.Point(12, 13);
            this.enableNegativeMarkingCheckBox.Name = "enableNegativeMarkingCheckBox";
            this.enableNegativeMarkingCheckBox.Size = new System.Drawing.Size(146, 17);
            this.enableNegativeMarkingCheckBox.TabIndex = 0;
            this.enableNegativeMarkingCheckBox.Text = "Enable Negative Marking";
            this.enableNegativeMarkingCheckBox.UseVisualStyleBackColor = true;
            // 
            // chosenFormFilenameLabel
            // 
            this.chosenFormFilenameLabel.AutoSize = true;
            this.chosenFormFilenameLabel.Location = new System.Drawing.Point(140, 35);
            this.chosenFormFilenameLabel.Name = "chosenFormFilenameLabel";
            this.chosenFormFilenameLabel.Size = new System.Drawing.Size(0, 13);
            this.chosenFormFilenameLabel.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(477, 320);
            this.Controls.Add(this.chosenFormFilenameLabel);
            this.Controls.Add(this.testSettingsPanel);
            this.Controls.Add(this.savePathTextBox);
            this.Controls.Add(this.chooseSavePathButton);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.chooseFormButton);
            this.Controls.Add(this.chooseFileLabel);
            this.Name = "Form1";
            this.Text = "Blackboard Test Creator";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.testSettingsPanel.ResumeLayout(false);
            this.testSettingsPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label chooseFileLabel;
        private System.Windows.Forms.Button chooseFormButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.Button chooseSavePathButton;
        private System.Windows.Forms.TextBox savePathTextBox;
        private System.Windows.Forms.Panel testSettingsPanel;
        private System.Windows.Forms.Label questionScoreLabel;
        private System.Windows.Forms.TextBox questionScoreTextBox;
        private System.Windows.Forms.CheckBox allowOverallNegativeScoreCheckBox;
        private System.Windows.Forms.CheckBox enableNegativeMarkingCheckBox;
        private System.Windows.Forms.Label chosenFormFilenameLabel;
    }
}

