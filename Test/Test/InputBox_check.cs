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
    public partial class InputBox_check : MetroFramework.Forms.MetroForm
    {
        public string oldItem = "";
        public string oldequipment = "";
        public string oldcheckresult = "";
        public string oldcheckdetail = "";
        public string oldcheckstrand = "";
        public string oldchecknote = "";
        public int style = 0;

        private string newdetail;
        private string newnote;
        private string newresult;
        private string newstrand;
       
        public InputBox_check()
        {
            InitializeComponent();
        }

        private void InputBox_check_Load(object sender, EventArgs e)
        {
            checkitem.Text = oldItem.ToString();
            checkequipment.Text = oldequipment.ToString();
            taskDetail.Text = oldcheckdetail.ToString();
            checkresulttxt.Text = oldcheckresult.ToString();
            strandtxt.Text = oldcheckstrand.ToString();
            notetxt.Text = oldchecknote.ToString();
            

            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);
        }
        
        public string Getdetail()
        {
            return newdetail;
        }
        public string Getnote()
        {
            return newnote;
        }
        public string Getresult()
        {
            return newresult;
        }
        public string Getstrand()
        {
            return newstrand;
        }

        private void metroLabel4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            newdetail = taskDetail.Text;
            newresult = checkresulttxt.Text;
            newstrand = strandtxt.Text;
            newnote = notetxt.Text;
        }
    }
}
