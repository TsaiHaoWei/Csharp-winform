namespace Test
{
    partial class InputBox_reject
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InputBox_reject));
            this.metroStyleManager1 = new MetroFramework.Components.MetroStyleManager(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.taskDetail = new MetroFramework.Controls.MetroTextBox();
            this.metroLabel3 = new MetroFramework.Controls.MetroLabel();
            this.taskitem = new MetroFramework.Controls.MetroLabel();
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.metroStyleManager1)).BeginInit();
            this.SuspendLayout();
            // 
            // metroStyleManager1
            // 
            this.metroStyleManager1.Owner = null;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.panel1.Location = new System.Drawing.Point(30, 139);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(574, 3);
            this.panel1.TabIndex = 42;
            // 
            // taskDetail
            // 
            this.taskDetail.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            // 
            // 
            // 
            this.taskDetail.CustomButton.Image = null;
            this.taskDetail.CustomButton.Location = new System.Drawing.Point(321, 2);
            this.taskDetail.CustomButton.Name = "";
            this.taskDetail.CustomButton.Size = new System.Drawing.Size(65, 65);
            this.taskDetail.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.taskDetail.CustomButton.TabIndex = 1;
            this.taskDetail.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.taskDetail.CustomButton.UseSelectable = true;
            this.taskDetail.CustomButton.Visible = false;
            this.taskDetail.Icon = ((System.Drawing.Image)(resources.GetObject("taskDetail.Icon")));
            this.taskDetail.IconRight = true;
            this.taskDetail.Lines = new string[0];
            this.taskDetail.Location = new System.Drawing.Point(195, 174);
            this.taskDetail.MaxLength = 32767;
            this.taskDetail.Multiline = true;
            this.taskDetail.Name = "taskDetail";
            this.taskDetail.PasswordChar = '\0';
            this.taskDetail.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.taskDetail.SelectedText = "";
            this.taskDetail.SelectionLength = 0;
            this.taskDetail.SelectionStart = 0;
            this.taskDetail.ShortcutsEnabled = true;
            this.taskDetail.Size = new System.Drawing.Size(389, 70);
            this.taskDetail.TabIndex = 39;
            this.taskDetail.UseSelectable = true;
            this.taskDetail.UseStyleColors = true;
            this.taskDetail.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.taskDetail.WaterMarkFont = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            // 
            // metroLabel3
            // 
            this.metroLabel3.AutoSize = true;
            this.metroLabel3.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabel3.Location = new System.Drawing.Point(30, 183);
            this.metroLabel3.Name = "metroLabel3";
            this.metroLabel3.Size = new System.Drawing.Size(92, 25);
            this.metroLabel3.TabIndex = 36;
            this.metroLabel3.Text = "退文原因:";
            this.metroLabel3.UseStyleColors = true;
            // 
            // taskitem
            // 
            this.taskitem.AutoSize = true;
            this.taskitem.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.taskitem.Location = new System.Drawing.Point(219, 102);
            this.taskitem.Name = "taskitem";
            this.taskitem.Size = new System.Drawing.Size(18, 25);
            this.taskitem.TabIndex = 37;
            this.taskitem.Text = "1";
            this.taskitem.UseStyleColors = true;
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabel1.Location = new System.Drawing.Point(30, 102);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(130, 25);
            this.metroLabel1.TabIndex = 38;
            this.metroLabel1.Text = "退文表單編號:";
            this.metroLabel1.UseStyleColors = true;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("新細明體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(151, 260);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(86, 37);
            this.button1.TabIndex = 35;
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
            this.button2.Location = new System.Drawing.Point(336, 260);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(92, 37);
            this.button2.TabIndex = 34;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = false;
            // 
            // InputBox_reject
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(621, 328);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.taskDetail);
            this.Controls.Add(this.metroLabel3);
            this.Controls.Add(this.taskitem);
            this.Controls.Add(this.metroLabel1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "InputBox_reject";
            this.Text = "InputBox_reject";
            this.Load += new System.EventHandler(this.InputBox_reject_Load);
            ((System.ComponentModel.ISupportInitialize)(this.metroStyleManager1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MetroFramework.Components.MetroStyleManager metroStyleManager1;
        private System.Windows.Forms.Panel panel1;
        private MetroFramework.Controls.MetroTextBox taskDetail;
        private MetroFramework.Controls.MetroLabel metroLabel3;
        private MetroFramework.Controls.MetroLabel taskitem;
        private MetroFramework.Controls.MetroLabel metroLabel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}