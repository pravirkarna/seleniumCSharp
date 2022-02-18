namespace MaximTrak.Automation.TestSuiteExecutor
{
    partial class AppSetter
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

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.combBowser = new System.Windows.Forms.ComboBox();
            this.lblBrowser = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtDegreeOfParall = new System.Windows.Forms.TextBox();
            this.lblDegreeOfParallel = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.lblPassword = new System.Windows.Forms.Label();
            this.txtSendMailFrom = new System.Windows.Forms.TextBox();
            this.lblSendMailFrom = new System.Windows.Forms.Label();
            this.txtSendEmailTo = new System.Windows.Forms.TextBox();
            this.lblSendEmailTo = new System.Windows.Forms.Label();
            this.combSendMail = new System.Windows.Forms.ComboBox();
            this.lblSendMail = new System.Windows.Forms.Label();
            this.combLoadTCFrmFile = new System.Windows.Forms.ComboBox();
            this.lblLoadTCFrmFile = new System.Windows.Forms.Label();
            this.combTestDataFiles = new System.Windows.Forms.ComboBox();
            this.lblTestDataFile = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.lblAppNotification = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // combBowser
            // 
            this.combBowser.FormattingEnabled = true;
            this.combBowser.Items.AddRange(new object[] {
            "Chrome",
            "Firefox",
            "IE"});
            this.combBowser.Location = new System.Drawing.Point(383, 8);
            this.combBowser.Name = "combBowser";
            this.combBowser.Size = new System.Drawing.Size(157, 21);
            this.combBowser.TabIndex = 2;
            // 
            // lblBrowser
            // 
            this.lblBrowser.AutoSize = true;
            this.lblBrowser.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBrowser.Location = new System.Drawing.Point(261, 11);
            this.lblBrowser.Name = "lblBrowser";
            this.lblBrowser.Size = new System.Drawing.Size(52, 13);
            this.lblBrowser.TabIndex = 3;
            this.lblBrowser.Text = "Browser";
            this.lblBrowser.Click += new System.EventHandler(this.label2_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.txtDegreeOfParall);
            this.panel1.Controls.Add(this.lblDegreeOfParallel);
            this.panel1.Controls.Add(this.txtPassword);
            this.panel1.Controls.Add(this.lblPassword);
            this.panel1.Controls.Add(this.txtSendMailFrom);
            this.panel1.Controls.Add(this.lblSendMailFrom);
            this.panel1.Controls.Add(this.txtSendEmailTo);
            this.panel1.Controls.Add(this.lblSendEmailTo);
            this.panel1.Controls.Add(this.combSendMail);
            this.panel1.Controls.Add(this.lblSendMail);
            this.panel1.Controls.Add(this.combLoadTCFrmFile);
            this.panel1.Controls.Add(this.lblLoadTCFrmFile);
            this.panel1.Controls.Add(this.combTestDataFiles);
            this.panel1.Controls.Add(this.lblTestDataFile);
            this.panel1.Controls.Add(this.lblBrowser);
            this.panel1.Controls.Add(this.combBowser);
            this.panel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(549, 200);
            this.panel1.TabIndex = 4;
            // 
            // txtDegreeOfParall
            // 
            this.txtDegreeOfParall.Location = new System.Drawing.Point(101, 11);
            this.txtDegreeOfParall.Name = "txtDegreeOfParall";
            this.txtDegreeOfParall.Size = new System.Drawing.Size(47, 20);
            this.txtDegreeOfParall.TabIndex = 19;
            // 
            // lblDegreeOfParallel
            // 
            this.lblDegreeOfParallel.AutoSize = true;
            this.lblDegreeOfParallel.Location = new System.Drawing.Point(5, 11);
            this.lblDegreeOfParallel.Name = "lblDegreeOfParallel";
            this.lblDegreeOfParallel.Size = new System.Drawing.Size(93, 13);
            this.lblDegreeOfParallel.TabIndex = 18;
            this.lblDegreeOfParallel.Text = "DegreeOfParall";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(101, 108);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(144, 20);
            this.txtPassword.TabIndex = 17;
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(5, 111);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(61, 13);
            this.lblPassword.TabIndex = 16;
            this.lblPassword.Text = "Password";
            this.lblPassword.Click += new System.EventHandler(this.label9_Click);
            // 
            // txtSendMailFrom
            // 
            this.txtSendMailFrom.Location = new System.Drawing.Point(383, 91);
            this.txtSendMailFrom.Name = "txtSendMailFrom";
            this.txtSendMailFrom.Size = new System.Drawing.Size(157, 20);
            this.txtSendMailFrom.TabIndex = 15;
            // 
            // lblSendMailFrom
            // 
            this.lblSendMailFrom.AutoSize = true;
            this.lblSendMailFrom.Location = new System.Drawing.Point(261, 91);
            this.lblSendMailFrom.Name = "lblSendMailFrom";
            this.lblSendMailFrom.Size = new System.Drawing.Size(86, 13);
            this.lblSendMailFrom.TabIndex = 14;
            this.lblSendMailFrom.Text = "SendMailFrom";
            // 
            // txtSendEmailTo
            // 
            this.txtSendEmailTo.Location = new System.Drawing.Point(101, 79);
            this.txtSendEmailTo.Name = "txtSendEmailTo";
            this.txtSendEmailTo.Size = new System.Drawing.Size(144, 20);
            this.txtSendEmailTo.TabIndex = 13;
            // 
            // lblSendEmailTo
            // 
            this.lblSendEmailTo.AutoSize = true;
            this.lblSendEmailTo.Location = new System.Drawing.Point(5, 79);
            this.lblSendEmailTo.Name = "lblSendEmailTo";
            this.lblSendEmailTo.Size = new System.Drawing.Size(81, 13);
            this.lblSendEmailTo.TabIndex = 12;
            this.lblSendEmailTo.Text = "SendEmailTo";
            // 
            // combSendMail
            // 
            this.combSendMail.FormattingEnabled = true;
            this.combSendMail.Items.AddRange(new object[] {
            "YES",
            "NO"});
            this.combSendMail.Location = new System.Drawing.Point(383, 45);
            this.combSendMail.Name = "combSendMail";
            this.combSendMail.Size = new System.Drawing.Size(157, 21);
            this.combSendMail.TabIndex = 11;
            // 
            // lblSendMail
            // 
            this.lblSendMail.AutoSize = true;
            this.lblSendMail.Location = new System.Drawing.Point(261, 48);
            this.lblSendMail.Name = "lblSendMail";
            this.lblSendMail.Size = new System.Drawing.Size(59, 13);
            this.lblSendMail.TabIndex = 10;
            this.lblSendMail.Text = "SendMail";
            // 
            // combLoadTCFrmFile
            // 
            this.combLoadTCFrmFile.FormattingEnabled = true;
            this.combLoadTCFrmFile.Items.AddRange(new object[] {
            "YES",
            "NO"});
            this.combLoadTCFrmFile.Location = new System.Drawing.Point(383, 127);
            this.combLoadTCFrmFile.Name = "combLoadTCFrmFile";
            this.combLoadTCFrmFile.Size = new System.Drawing.Size(157, 21);
            this.combLoadTCFrmFile.TabIndex = 9;
            // 
            // lblLoadTCFrmFile
            // 
            this.lblLoadTCFrmFile.AutoSize = true;
            this.lblLoadTCFrmFile.Location = new System.Drawing.Point(261, 130);
            this.lblLoadTCFrmFile.Name = "lblLoadTCFrmFile";
            this.lblLoadTCFrmFile.Size = new System.Drawing.Size(109, 13);
            this.lblLoadTCFrmFile.TabIndex = 8;
            this.lblLoadTCFrmFile.Text = "LoadTCFrmExtFile";
            // 
            // combTestDataFiles
            // 
            this.combTestDataFiles.FormattingEnabled = true;
            this.combTestDataFiles.Items.AddRange(new object[] {
            "QA",
            "RC",
            "UAT",
            "PROD"});
            this.combTestDataFiles.Location = new System.Drawing.Point(101, 45);
            this.combTestDataFiles.Name = "combTestDataFiles";
            this.combTestDataFiles.Size = new System.Drawing.Size(144, 21);
            this.combTestDataFiles.TabIndex = 7;
            // 
            // lblTestDataFile
            // 
            this.lblTestDataFile.AutoSize = true;
            this.lblTestDataFile.Location = new System.Drawing.Point(5, 45);
            this.lblTestDataFile.Name = "lblTestDataFile";
            this.lblTestDataFile.Size = new System.Drawing.Size(85, 13);
            this.lblTestDataFile.TabIndex = 6;
            this.lblTestDataFile.Text = "TestDataFiles";
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.btnClose);
            this.panel2.Controls.Add(this.btnClear);
            this.panel2.Controls.Add(this.btnSave);
            this.panel2.Location = new System.Drawing.Point(12, 218);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(286, 44);
            this.panel2.TabIndex = 5;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(193, 9);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(85, 27);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(100, 9);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(87, 27);
            this.btnClear.TabIndex = 1;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(7, 9);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(87, 27);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblAppNotification
            // 
            this.lblAppNotification.AutoSize = true;
            this.lblAppNotification.Location = new System.Drawing.Point(12, 273);
            this.lblAppNotification.Name = "lblAppNotification";
            this.lblAppNotification.Size = new System.Drawing.Size(0, 13);
            this.lblAppNotification.TabIndex = 6;
            // 
            // AppSetter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(573, 295);
            this.Controls.Add(this.lblAppNotification);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "AppSetter";
            this.Text = "App";
            this.Load += new System.EventHandler(this.AppSetter_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.ComboBox combBowser;
        private System.Windows.Forms.Label lblBrowser;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox combTestDataFiles;
        private System.Windows.Forms.Label lblTestDataFile;
        private System.Windows.Forms.ComboBox combLoadTCFrmFile;
        private System.Windows.Forms.Label lblLoadTCFrmFile;
        private System.Windows.Forms.ComboBox combSendMail;
        private System.Windows.Forms.Label lblSendMail;
        private System.Windows.Forms.Label lblSendEmailTo;
        private System.Windows.Forms.TextBox txtSendEmailTo;
        private System.Windows.Forms.TextBox txtSendMailFrom;
        private System.Windows.Forms.Label lblSendMailFrom;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtDegreeOfParall;
        private System.Windows.Forms.Label lblDegreeOfParallel;
        private System.Windows.Forms.Label lblAppNotification;
    }
}