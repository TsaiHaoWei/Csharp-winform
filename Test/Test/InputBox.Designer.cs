namespace Test
{
    partial class InputBox
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InputBox));
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.taskitem = new MetroFramework.Controls.MetroLabel();
            this.metroLabel3 = new MetroFramework.Controls.MetroLabel();
            this.taskDetail = new MetroFramework.Controls.MetroTextBox();
            this.tasknote = new MetroFramework.Controls.MetroTextBox();
            this.metroLabel4 = new MetroFramework.Controls.MetroLabel();
            this.metroStyleManager1 = new MetroFramework.Components.MetroStyleManager(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.metroStyleManager1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(199, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(158, 32);
            this.label1.TabIndex = 2;
            this.label1.Text = "資料更改!";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(220, 431);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(86, 37);
            this.button1.TabIndex = 3;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button2.ForeColor = System.Drawing.Color.White;
            this.button2.Location = new System.Drawing.Point(405, 431);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(92, 37);
            this.button2.TabIndex = 3;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(77, 57);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(79, 68);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabel1.Location = new System.Drawing.Point(77, 161);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(168, 25);
            this.metroLabel1.TabIndex = 18;
            this.metroLabel1.Text = "工作任務表單編號:";
            this.metroLabel1.UseStyleColors = true;
            // 
            // taskitem
            // 
            this.taskitem.AutoSize = true;
            this.taskitem.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.taskitem.Location = new System.Drawing.Point(266, 161);
            this.taskitem.Name = "taskitem";
            this.taskitem.Size = new System.Drawing.Size(18, 25);
            this.taskitem.TabIndex = 18;
            this.taskitem.Text = "1";
            this.taskitem.UseStyleColors = true;
            // 
            // metroLabel3
            // 
            this.metroLabel3.AutoSize = true;
            this.metroLabel3.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabel3.Location = new System.Drawing.Point(77, 242);
            this.metroLabel3.Name = "metroLabel3";
            this.metroLabel3.Size = new System.Drawing.Size(92, 25);
            this.metroLabel3.TabIndex = 18;
            this.metroLabel3.Text = "工作內容:";
            this.metroLabel3.UseStyleColors = true;
            // 
            // taskDetail
            // 
            this.taskDetail.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            // 
            // 
            // 
            this.taskDetail.CustomButton.Image = null;
            this.taskDetail.CustomButton.Location = new System.Drawing.Point(295, 1);
            this.taskDetail.CustomButton.Name = "";
            this.taskDetail.CustomButton.Size = new System.Drawing.Size(93, 93);
            this.taskDetail.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.taskDetail.CustomButton.TabIndex = 1;
            this.taskDetail.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.taskDetail.CustomButton.UseSelectable = true;
            this.taskDetail.CustomButton.Visible = false;
            this.taskDetail.Icon = ((System.Drawing.Image)(resources.GetObject("taskDetail.Icon")));
            this.taskDetail.IconRight = true;
            this.taskDetail.Lines = new string[0];
            this.taskDetail.Location = new System.Drawing.Point(280, 242);
            this.taskDetail.MaxLength = 32767;
            this.taskDetail.Multiline = true;
            this.taskDetail.Name = "taskDetail";
            this.taskDetail.PasswordChar = '\0';
            this.taskDetail.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.taskDetail.SelectedText = "";
            this.taskDetail.SelectionLength = 0;
            this.taskDetail.SelectionStart = 0;
            this.taskDetail.ShortcutsEnabled = true;
            this.taskDetail.Size = new System.Drawing.Size(389, 95);
            this.taskDetail.TabIndex = 19;
            this.taskDetail.UseSelectable = true;
            this.taskDetail.UseStyleColors = true;
            this.taskDetail.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.taskDetail.WaterMarkFont = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            // 
            // tasknote
            // 
            this.tasknote.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            // 
            // 
            // 
            this.tasknote.CustomButton.Image = null;
            this.tasknote.CustomButton.Location = new System.Drawing.Point(357, 1);
            this.tasknote.CustomButton.Name = "";
            this.tasknote.CustomButton.Size = new System.Drawing.Size(31, 31);
            this.tasknote.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.tasknote.CustomButton.TabIndex = 1;
            this.tasknote.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.tasknote.CustomButton.UseSelectable = true;
            this.tasknote.CustomButton.Visible = false;
            this.tasknote.Icon = ((System.Drawing.Image)(resources.GetObject("tasknote.Icon")));
            this.tasknote.IconRight = true;
            this.tasknote.Lines = new string[0];
            this.tasknote.Location = new System.Drawing.Point(280, 360);
            this.tasknote.MaxLength = 32767;
            this.tasknote.Multiline = true;
            this.tasknote.Name = "tasknote";
            this.tasknote.PasswordChar = '\0';
            this.tasknote.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.tasknote.SelectedText = "";
            this.tasknote.SelectionLength = 0;
            this.tasknote.SelectionStart = 0;
            this.tasknote.ShortcutsEnabled = true;
            this.tasknote.Size = new System.Drawing.Size(389, 33);
            this.tasknote.TabIndex = 21;
            this.tasknote.UseSelectable = true;
            this.tasknote.UseStyleColors = true;
            this.tasknote.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.tasknote.WaterMarkFont = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            // 
            // metroLabel4
            // 
            this.metroLabel4.AutoSize = true;
            this.metroLabel4.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabel4.Location = new System.Drawing.Point(77, 360);
            this.metroLabel4.Name = "metroLabel4";
            this.metroLabel4.Size = new System.Drawing.Size(54, 25);
            this.metroLabel4.TabIndex = 20;
            this.metroLabel4.Text = "備註:";
            this.metroLabel4.UseStyleColors = true;
            // 
            // metroStyleManager1
            // 
            this.metroStyleManager1.Owner = this;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.panel1.Location = new System.Drawing.Point(77, 201);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(592, 3);
            this.panel1.TabIndex = 33;
            // 
            // InputBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(701, 510);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.tasknote);
            this.Controls.Add(this.metroLabel4);
            this.Controls.Add(this.taskDetail);
            this.Controls.Add(this.metroLabel3);
            this.Controls.Add(this.taskitem);
            this.Controls.Add(this.metroLabel1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "InputBox";
            this.Style = MetroFramework.MetroColorStyle.Default;
            this.Text = "TaskCard更改";
            this.Load += new System.EventHandler(this.InputBox_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.metroStyleManager1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private MetroFramework.Controls.MetroLabel metroLabel1;
        private MetroFramework.Controls.MetroLabel taskitem;
        private MetroFramework.Controls.MetroLabel metroLabel3;
        private MetroFramework.Controls.MetroTextBox taskDetail;
        private MetroFramework.Controls.MetroTextBox tasknote;
        private MetroFramework.Controls.MetroLabel metroLabel4;
        private MetroFramework.Components.MetroStyleManager metroStyleManager1;
        private System.Windows.Forms.Panel panel1;
    }
}