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
    public partial class InputBox_reject : MetroFramework.Forms.MetroForm
    {
        public string taskcode = "";
    
        public int style = 0;

        private string reson;
       

        public InputBox_reject()
        {
            InitializeComponent();
        }
        public string Getreson()
        {
            return reson;
        }

        private void InputBox_reject_Load(object sender, EventArgs e)
        {
            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);
            taskitem.Text = taskcode.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            reson = taskDetail.Text;
        }
    }
}
