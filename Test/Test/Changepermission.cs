using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
    public partial class Changepermission : MetroFramework.Forms.MetroForm
    {
        public int style = 0;
        public int permiss = 0;
        public string name = "";
        public string team = "";
        public string position = "";

        private int reper = 0;     
             

        public Changepermission()
        {
            InitializeComponent();
        }

        private void Changepermission_Load(object sender, EventArgs e)
        {
            name1.Text = this.name;
            team1.Text = this.team;
            position1.Text = this.position;
            oldpermission.Text = this.permiss.ToString();

            
            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);
            //changeuser.Text = Name;
        }

        private void save_Click(object sender, EventArgs e)//存取
        {
            try { reper = Int32.Parse(newpermission.Text); }
            catch (Exception ex) {Console.WriteLine(ex.ToString()); }
        }
        public int Getper()
        {
            return reper;
        }

        private void cancel_Click(object sender, EventArgs e)//取消
        {

        }
    }
}
