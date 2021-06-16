using MySql.Data.MySqlClient;
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
    public partial class Reject : MetroFramework.Forms.MetroForm
    {
        public string IP = "";
        public string port = "";
        public int style = 0;
        public string taskname = "";
        public string Name = "";
        public Reject()
        {
            InitializeComponent();
           
        }

        private void Reject_Load(object sender, EventArgs e)
        {
            this.style = style;
            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);

            string connString = "server=" + IP.ToString() + ";port=" + port.ToString() + ";user id="+Name.ToString()+";password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫
            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

            //有表格
            string sql = @"select * from `taskrejected` WHERE TaskCard = @taskname ";

            MySqlCommand cmdusd = new MySqlCommand(sql, conn);
            cmdusd.Parameters.AddWithValue("@taskname", taskname);

            MySqlDataReader drusd = cmdusd.ExecuteReader();

            while (drusd.Read())
            {
             
                string reason= drusd["Reason"].ToString();
                string state = drusd["State"].ToString();
                string code = drusd["TaskCardCode"].ToString();
                string engineer = drusd["Engineer"].ToString();
                string Time = drusd["UpdateTime"].ToString();

                Gridsch.Rows.Add(new Object[] {code,reason,Time,engineer,state});
                
            }
            metroLabel1.Text = taskname;
            drusd.Close();
            cmdusd.Dispose();
            conn.Dispose();
           
        }
    }
}
