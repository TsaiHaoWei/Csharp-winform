namespace Test
{
    partial class Changeaccount
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Changeaccount));
            this.metroStyleManager1 = new MetroFramework.Components.MetroStyleManager(this.components);
            this.metroLabUser = new MetroFramework.Controls.MetroLabel();
            this.changeuser = new MetroFramework.Controls.MetroLabel();
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.metroLabel2 = new MetroFramework.Controls.MetroLabel();
            this.oldsecret = new MetroFramework.Controls.MetroTextBox();
            this.newsecret = new MetroFramework.Controls.MetroTextBox();
            this.metroButton1 = new MetroFramework.Controls.MetroButton();
            ((System.ComponentModel.ISupportInitialize)(this.metroStyleManager1)).BeginInit();
            this.SuspendLayout();
            // 
            // metroStyleManager1
            // 
            this.metroStyleManager1.Owner = this;
            // 
            // metroLabUser
            // 
            this.metroLabUser.AutoSize = true;
            this.metroLabUser.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabUser.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            this.metroLabUser.Location = new System.Drawing.Point(101, 92);
            this.metroLabUser.Name = "metroLabUser";
            this.metroLabUser.Size = new System.Drawing.Size(112, 25);
            this.metroLabUser.TabIndex = 0;
            this.metroLabUser.Text = "使用者帳號";
            this.metroLabUser.UseStyleColors = true;
            // 
            // changeuser
            // 
            this.changeuser.AutoSize = true;
            this.changeuser.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.changeuser.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            this.changeuser.Location = new System.Drawing.Point(268, 92);
            this.changeuser.Name = "changeuser";
            this.changeuser.Size = new System.Drawing.Size(59, 25);
            this.changeuser.TabIndex = 0;
            this.changeuser.Text = "name";
            this.changeuser.UseStyleColors = true;
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabel1.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            this.metroLabel1.Location = new System.Drawing.Point(101, 153);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(112, 25);
            this.metroLabel1.TabIndex = 0;
            this.metroLabel1.Text = "舊密碼確認";
            this.metroLabel1.UseStyleColors = true;
            // 
            // metroLabel2
            // 
            this.metroLabel2.AutoSize = true;
            this.metroLabel2.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabel2.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            this.metroLabel2.Location = new System.Drawing.Point(101, 206);
            this.metroLabel2.Name = "metroLabel2";
            this.metroLabel2.Size = new System.Drawing.Size(112, 25);
            this.metroLabel2.TabIndex = 0;
            this.metroLabel2.Text = "新密碼確認";
            this.metroLabel2.UseStyleColors = true;
            // 
            // oldsecret
            // 
            // 
            // 
            // 
            this.oldsecret.CustomButton.Image = null;
            this.oldsecret.CustomButton.Location = new System.Drawing.Point(116, 1);
            this.oldsecret.CustomButton.Name = "";
            this.oldsecret.CustomButton.Size = new System.Drawing.Size(21, 21);
            this.oldsecret.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.oldsecret.CustomButton.TabIndex = 1;
            this.oldsecret.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.oldsecret.CustomButton.UseSelectable = true;
            this.oldsecret.CustomButton.Visible = false;
            this.oldsecret.DisplayIcon = true;
            this.oldsecret.Icon = ((System.Drawing.Image)(resources.GetObject("oldsecret.Icon")));
            this.oldsecret.Lines = new string[0];
            this.oldsecret.Location = new System.Drawing.Point(252, 153);
            this.oldsecret.MaxLength = 32767;
            this.oldsecret.Name = "oldsecret";
            this.oldsecret.PasswordChar = '●';
            this.oldsecret.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.oldsecret.SelectedText = "";
            this.oldsecret.SelectionLength = 0;
            this.oldsecret.SelectionStart = 0;
            this.oldsecret.ShortcutsEnabled = true;
            this.oldsecret.Size = new System.Drawing.Size(138, 23);
            this.oldsecret.TabIndex = 1;
            this.oldsecret.UseSelectable = true;
            this.oldsecret.UseStyleColors = true;
            this.oldsecret.UseSystemPasswordChar = true;
            this.oldsecret.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.oldsecret.WaterMarkFont = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel);
            // 
            // newsecret
            // 
            // 
            // 
            // 
            this.newsecret.CustomButton.Image = null;
            this.newsecret.CustomButton.Location = new System.Drawing.Point(116, 1);
            this.newsecret.CustomButton.Name = "";
            this.newsecret.CustomButton.Size = new System.Drawing.Size(21, 21);
            this.newsecret.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.newsecret.CustomButton.TabIndex = 1;
            this.newsecret.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.newsecret.CustomButton.UseSelectable = true;
            this.newsecret.CustomButton.Visible = false;
            this.newsecret.DisplayIcon = true;
            this.newsecret.Icon = ((System.Drawing.Image)(resources.GetObject("newsecret.Icon")));
            this.newsecret.Lines = new string[0];
            this.newsecret.Location = new System.Drawing.Point(252, 206);
            this.newsecret.MaxLength = 32767;
            this.newsecret.Name = "newsecret";
            this.newsecret.PasswordChar = '●';
            this.newsecret.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.newsecret.SelectedText = "";
            this.newsecret.SelectionLength = 0;
            this.newsecret.SelectionStart = 0;
            this.newsecret.ShortcutsEnabled = true;
            this.newsecret.Size = new System.Drawing.Size(138, 23);
            this.newsecret.TabIndex = 2;
            this.newsecret.UseSelectable = true;
            this.newsecret.UseStyleColors = true;
            this.newsecret.UseSystemPasswordChar = true;
            this.newsecret.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.newsecret.WaterMarkFont = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel);
            // 
            // metroButton1
            // 
            this.metroButton1.BackColor = System.Drawing.Color.White;
            this.metroButton1.FontSize = MetroFramework.MetroButtonSize.Tall;
            this.metroButton1.Highlight = true;
            this.metroButton1.Location = new System.Drawing.Point(374, 247);
            this.metroButton1.Name = "metroButton1";
            this.metroButton1.Size = new System.Drawing.Size(101, 26);
            this.metroButton1.TabIndex = 3;
            this.metroButton1.Text = "變更密碼";
            this.metroButton1.UseCustomBackColor = true;
            this.metroButton1.UseSelectable = true;
            this.metroButton1.UseStyleColors = true;
            this.metroButton1.Click += new System.EventHandler(this.metroButton1_Click);
            // 
            // Changeaccount
            // 
            this.AcceptButton = this.metroButton1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(498, 296);
            this.Controls.Add(this.metroButton1);
            this.Controls.Add(this.newsecret);
            this.Controls.Add(this.oldsecret);
            this.Controls.Add(this.changeuser);
            this.Controls.Add(this.metroLabel2);
            this.Controls.Add(this.metroLabel1);
            this.Controls.Add(this.metroLabUser);
            this.Name = "Changeaccount";
            this.Text = "Changeaccount";
            this.Load += new System.EventHandler(this.Changeaccount_Load);
            ((System.ComponentModel.ISupportInitialize)(this.metroStyleManager1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MetroFramework.Components.MetroStyleManager metroStyleManager1;
        private MetroFramework.Controls.MetroLabel changeuser;
        private MetroFramework.Controls.MetroLabel metroLabUser;
        private MetroFramework.Controls.MetroTextBox newsecret;
        private MetroFramework.Controls.MetroTextBox oldsecret;
        private MetroFramework.Controls.MetroLabel metroLabel2;
        private MetroFramework.Controls.MetroLabel metroLabel1;
        private MetroFramework.Controls.MetroButton metroButton1;
    }
}