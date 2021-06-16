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
    public partial class InputBox : MetroFramework.Forms.MetroForm
    {
        public string oldname = "";
        public string oldetail = "";
        public string oldnote = "";
        public int style = 0;
        string rename = "";
        public InputBox()
        {
            InitializeComponent();
      
        }
        private string newdetail;
        private string newnote;


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void InputBox_Load(object sender, EventArgs e)
        {
            taskDetail.Text = oldetail.ToString();
            tasknote.Text = oldnote.ToString();
            taskitem.Text = oldname.ToString();
            rename = Name.ToString();

            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            newdetail = taskDetail.Text;
            newnote = tasknote.Text;
        }
        public string Getdetail()
        {
            return newdetail;
        }
        public string Getnote()
        {
            return newnote;
        }



        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
