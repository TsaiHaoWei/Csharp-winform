using MySql.Data.MySqlClient;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

using Spire.Doc;
using Spire.Doc.Interface;
using Spire.Doc.Documents;
using Spire.Doc.Utilities;
using Spire.Doc.Collections;

namespace Test
{
    public partial class ReadTask : MetroFramework.Forms.MetroForm
    {
        public string IP = "";
        public string port = "";
        public string Code = ""; //顯示此資料用的
        public string nametalk = "";
        public string accountname = "";//人名
        public string location = "";
        public string TaskCodeonly = "";
        public string Name = "";

 
        public int style = 0;
        string date = "";
        string Month = "";
        string[] photoname = { };//紀錄圖片路徑
        int photousecount = 0;//使用第幾張
        int photocount = 0;//紀錄圖片張數
        string photodate = "";
        string sign_save = "";
        string timeend = "";
        //防止資料一直新增
        int counttask = 1;
        int countcheck = 1;
        int countrecord = 1;

        int ena = 1;
        //匯出WORD的使用的
      
        private object Nothing = Missing.Value;//預設值
        private object IsReadOnly = false;//不僅僅可讀
        private Microsoft.Office.Interop.Word._Application wordApp;//Word應用程式
        private Microsoft.Office.Interop.Word._Document wordDoc;//Word文件
        private int tableCount;//Word表格數目(完整的)

        string Filename = "";//抓搜尋名字
        string connString = "";
        public ReadTask()
        {
           
            InitializeComponent();
        }

        private void ReadTask_Load(object sender, EventArgs e)
        {

            this.Size = new System.Drawing.Size(Screen.PrimaryScreen.WorkingArea.Width - 300, Screen.PrimaryScreen.WorkingArea.Height - 300);
            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);
            TaskCode.Text = Code;
//taskCard
            if (counttask == 1)
            {
                Taskgrid.Visible = true;
               Checkgrid.Visible = false;
                Recordgrid.Visible = false;
               
                int count = 1;
                connString = "server=" + IP.ToString() + ";port=" + port.ToString() + ";user id="+Name.ToString()+";password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

                //有表格
                string sql = @"select * from `" + Code.ToString().ToLower() + "` ";
                
                MySqlCommand cmdusd = new MySqlCommand(sql, conn);

                MySqlDataReader drusd = cmdusd.ExecuteReader();
                
                while (drusd.Read())
                {
                    //sqltaskname = drusd["TaskCardName"].ToString();
                    //if (count == 1) TaskName.Text += ":" + sqltaskname.ToString();
                   nametalk = drusd["TaskCardName"].ToString();
                    string item = drusd["Items"].ToString();
                    string detail = drusd["TaskDetail"].ToString();
                    string standard = drusd["StandardValue"].ToString();
                    string note = drusd["Note"].ToString();
                    Taskgrid.Rows.Add(new Object[] { item, detail, standard,note });

                }
                drusd.Close();
                cmdusd.Dispose();

                counttask++;//防止重複輸入
            }
            else
            {
                Taskgrid.Visible = true;
                Checkgrid.Visible = false ;
                Recordgrid.Visible = false;
              
            }
           

        }

        private void Taskbut_Click(object sender, EventArgs e)
        {//TASKCARD

            if (counttask == 1)
            {
               
                outbut.Visible = false;
                //addtask.Visible = true;
                //  addcheck.Visible = false;
                recordcombo.Visible = false;
                Taskgrid.Visible = true;
                Checkgrid.Visible = false;
                Recordgrid.Visible = false;

                int count = 1;
            
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

                //有表格
                string sql = @"select * from `" + Code.ToString().ToLower() + "`";

                MySqlCommand cmdusd = new MySqlCommand(sql, conn);
                MySqlDataReader drusd = cmdusd.ExecuteReader();
                while (drusd.Read())
                {
                    //sqltaskname = drusd["TaskCardName"].ToString();
                    //if (count == 1) TaskName.Text += ":" + sqltaskname.ToString();

                    string item = drusd["Items"].ToString();
                    string detail = drusd["TaskDetail"].ToString();
                    string standard = drusd["StandardValue"].ToString();
                    string note = drusd["Note"].ToString();
                    Taskgrid.Rows.Add(new Object[] { item, detail, standard, note });

                }
                drusd.Close();
                cmdusd.Dispose();

                counttask++;//防止重複輸入
            }
            else
            {
                
                outbut.Visible = false;
                // addtask.Visible = true;
                //addcheck.Visible = false;
                recordcombo.Visible = false;
                Taskgrid.Visible = true;
                Checkgrid.Visible = false;
                Recordgrid.Visible = false;

            }
        }

        private void Checkbut_Click(object sender, EventArgs e)
        {
            if (countcheck == 1)
            {
               
                outbut.Visible = false;
                // addtask.Visible = false;
                // addcheck.Visible = true;
                recordcombo.Visible = false;
                Taskgrid.Visible = false;
                Checkgrid.Visible = true;
                Recordgrid.Visible = false;


                
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

                //有表格
                string sql = @"select * from `" + Code.ToString().ToLower() + "-f1`";

                MySqlCommand cmdusd = new MySqlCommand(sql, conn);
                MySqlDataReader drusd = cmdusd.ExecuteReader();
                while (drusd.Read())
                {
                    string item = drusd["Items"].ToString();
                    string equip = drusd["Equipment"].ToString();
                    string detail = drusd["TaskDetail"].ToString();
                    string checkResult = drusd["CheckResult"].ToString();
                    string standard = drusd["StandardValue"].ToString();
                    string note = drusd["Note"].ToString();
                    Checkgrid.Rows.Add(new Object[] { item, equip, checkResult, detail, standard,note });

                }
                drusd.Close();
                cmdusd.Dispose();
                countcheck++;
            }
            else
            {
                
                outbut.Visible = false;
                //addtask.Visible = false;
                // addcheck.Visible = true;
                recordcombo.Visible = false;

                Taskgrid.Visible = false;
                Checkgrid.Visible = true;
                Recordgrid.Visible = false;
            }

        }

        private void recordbut_Click(object sender, EventArgs e)
        {
            //RECORD
            if (countrecord == 1)
            {
                
                recordcombo.Visible = true;
                Taskgrid.Visible = false;
                Checkgrid.Visible = false;
                Recordgrid.Visible = true;
          
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                                                                    //有表格
                                                                    //String select = @"select Name from `account` where Name = @user";
                //string sql = @"select * from `" + Code.ToString().ToLower() + "-f1-record` where TaskCardCode=@user";
                string sql = @"select * from `" + Code.ToString().ToLower() + "-f1-record` ";
                MySqlCommand cmdusd = new MySqlCommand(sql, conn);
                //cmdusd.Parameters.Add(new MySqlParameter("@user", TaskCodeonly.ToString()));
                MySqlDataReader drusd = cmdusd.ExecuteReader();
                string oldtime = "";
                while (drusd.Read())
                {
                    string time = drusd["FinishTime"].ToString();
                    char[] delimiterChars = { '\n', '\t', '\r', '\a', ':' };//remove \r\a
                    string[] datetest = time.Split(delimiterChars);
                   time = datetest[0] + ":" + datetest[1];
                    if (!oldtime.ToString().Equals(time.ToString()))
                    {                  
                        recordcombo.Items.Add(time.ToString());
                    }
                    oldtime = time;
                }                drusd.Close();
                cmdusd.Dispose();
                countrecord++;


            }
            else
            {
              
                recordcombo.Visible = true;
                Taskgrid.Visible = false;
                Checkgrid.Visible = false;
                Recordgrid.Visible = true;
            }

        }

        private void Taskgrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            string Name = Taskgrid.Rows[e.RowIndex].Cells[1].Value.ToString();
            string Item = Taskgrid.Rows[e.RowIndex].Cells[0].Value.ToString();
            string stand = Taskgrid.Rows[e.RowIndex].Cells[2].Value.ToString();
            string note = Taskgrid.Rows[e.RowIndex].Cells[3].Value.ToString();

            InputBox input = new InputBox();
            input.oldname = Item.ToString();
            input.oldetail = Name.ToString();
            input.oldnote = note.ToString();
            input.style = this.style;
            DialogResult dr = input.ShowDialog();
            if (dr == DialogResult.OK)
            {
                Taskgrid.Rows[e.RowIndex].Cells[1].Value = input.Getdetail();
                Taskgrid.Rows[e.RowIndex].Cells[3].Value = input.Getnote();
                //資料庫
             
                MySqlConnection myConnection = new MySqlConnection(connString);//連線資料庫

                //開啟資料庫
                myConnection.Open();
                for (int i = 0; i < Taskgrid.RowCount - 1; i++)
                {

                    // MessageBox.Show(Code.ToString().ToLower());
                    string inser = "UPDATE `" + Code.ToString().ToLower() + @"` SET TaskDetail = @detial,Note = @note WHERE Items= @item ";

                    using (MySqlCommand cmd = new MySqlCommand(inser, myConnection))
                    {
                        string tcdetail = Taskgrid.Rows[i].Cells[1].Value.ToString();
                        string tcitem = Taskgrid.Rows[i].Cells[0].Value.ToString();
                        string tcnote = Taskgrid.Rows[i].Cells[3].Value.ToString();

                        cmd.Parameters.AddWithValue("@detial", tcdetail.ToString());
                        cmd.Parameters.AddWithValue("@item", tcitem.ToString());
                        cmd.Parameters.AddWithValue("@note", tcnote.ToString());
                       
                            int index = cmd.ExecuteNonQuery();
                            cmd.Dispose();
                                   

                        

                    }

                }


                //關閉資料庫
                myConnection.Close();
                }
            else if (dr == DialogResult.Cancel)
            {
                //MessageBox.Show(input.GetchanMsg());
            }
        }

        private void Checkgrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        
            string equipment = Checkgrid.Rows[e.RowIndex].Cells[1].Value.ToString();
            string strand= Checkgrid.Rows[e.RowIndex].Cells[4].Value.ToString();
            string Item = Checkgrid.Rows[e.RowIndex].Cells[0].Value.ToString();
            string checkresult = Checkgrid.Rows[e.RowIndex].Cells[2].Value.ToString();
            string checkdetail = Checkgrid.Rows[e.RowIndex].Cells[3].Value.ToString();
            string note = Checkgrid.Rows[e.RowIndex].Cells[5].Value.ToString();

            InputBox_check input = new InputBox_check();
            input.oldItem = Item.ToString();
            input.oldequipment = equipment.ToString();
            input.oldcheckresult = checkresult.ToString();
            input.oldcheckdetail = checkdetail.ToString();
            input.oldcheckstrand = strand.ToString();
            input.oldchecknote = note.ToString();
            input.style = this.style;
            DialogResult dr = input.ShowDialog();
            if (dr == DialogResult.OK)
            {
            
                MySqlConnection myConnection = new MySqlConnection(connString.ToString());//連線資料庫
                                //開啟資料庫
                myConnection.Open();
                string DELET = "truncate table `" + Code.ToString().ToLower() + "-f1`";
                  using (MySqlCommand cmd = new MySqlCommand(DELET, myConnection))
                {
                    int index = cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                    Checkgrid.Rows[e.RowIndex].Cells[4].Value = input.Getstrand();
                Checkgrid.Rows[e.RowIndex].Cells[2].Value = input.Getresult();
                Checkgrid.Rows[e.RowIndex].Cells[3].Value = input.Getdetail();
                Checkgrid.Rows[e.RowIndex].Cells[5].Value = input.Getnote();
                //資料庫

              

                for (int i = 0; i < Checkgrid.RowCount - 1; i++)
                {
                   
                                                                                                               //順訊不能顛倒
                    string sql = @"INSERT INTO `" + Code.ToString().ToLower() + "-f1" + @"`(`TaskCardCode`,`TaskCardName`,`Items`,`TaskDetail`,`Equipment`,`CheckResult`,`StandardValue`,`Note`) VALUES
                                                       (@test1,@test2,@test3,@test4,@test5,@test6,@test7,@test8) ";

                    using (MySqlCommand cmd = new MySqlCommand(sql, myConnection))
                    {
                        cmd.Parameters.AddWithValue("@test3", Checkgrid.Rows[i].Cells[0].Value.ToString());
                        cmd.Parameters.AddWithValue("@test5", Checkgrid.Rows[i].Cells[1].Value.ToString());

                        cmd.Parameters.AddWithValue("@test2", nametalk.ToString());
                        cmd.Parameters.AddWithValue("@test1", Code.ToString());
                        cmd.Parameters.AddWithValue("@test4", Checkgrid.Rows[i].Cells[3].Value.ToString());
                        cmd.Parameters.AddWithValue("@test6", Checkgrid.Rows[i].Cells[2].Value.ToString());
                        cmd.Parameters.AddWithValue("@test7", Checkgrid.Rows[i].Cells[4].Value.ToString());
                        cmd.Parameters.AddWithValue("@test8", Checkgrid.Rows[i].Cells[5].Value.ToString());
                   

                        int index = cmd.ExecuteNonQuery();
                        cmd.Dispose();




                    }

                }


                //關閉資料庫
                myConnection.Close();
            }
            else if (dr == DialogResult.Cancel)
            {
                //MessageBox.Show(input.GetchanMsg());
            }
        }

        private void Recordgrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            photocount = 0;
            photoname = new string[] { };//紀錄圖片路徑
            photousecount = 0;
            pictureBox1.Size = new Size(120,70);
            pictureBox1.Location =new Point(0,30);
            //圖片顯示
            int i = 0;
                 
               pictureBox1.Visible = true;
            foreach(string fname in System.IO.Directory.GetFileSystemEntries(@"D:\SCD\下載\PhotoResult"))
              {
       
                  //Console.WriteLine(fname);
                if (fname.ToString().Contains(photodate.ToString()) && fname.ToString().Contains(Code.ToString().ToLower()))
                { photocount++;
          


                    // 調整陣列的大小
                    System.Array.Resize(ref photoname, photoname.Length + 1);
                    // 指定新的陣列值
                    photoname[photoname.Length - 1] = fname.ToString();
                    
                    
                }
                //Console.WriteLine(photocount);
                i++;
              }
            if (photocount == 0)//沒有照片時
            { pictureBox1.Visible = false; }
            else
            {           
                pictureBox1.Image = Image.FromFile(photoname[0]);
                timer1.Enabled = true;
            }

        }


     

        private void recordcombo_DropDownClosed(object sender, EventArgs e)
        {if (!recordcombo.Text.Equals(""))
            {
                outbut.Visible = true;
                Recordgrid.Rows.Clear();

                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

                string sql = @"select * from `" + Code.ToString().ToLower() + "-f1-record`";
                MySqlCommand cmdusd = new MySqlCommand(sql, conn);
                MySqlDataReader drusd = cmdusd.ExecuteReader();

                while (drusd.Read())
                {//月份讀取
                    int read = 0;
                    //string Date = drusd["Time"].ToString().Remove(10);
                    //Date = Date.ToString().TrimEnd('上');
                    char[] delimiterChars = { '\n', '\t', '\r', '\a', '-', '/' };//remove \r\a
                                                                                 //string[] tests = Date.Split(delimiterChars);

                    string test = drusd["Name"].ToString();
                    string item = drusd["Items"].ToString();
                    string checkdetail = drusd["CheckDetail"].ToString();
                    string CheckResult = drusd["CheckResult"].ToString();
                    string result = drusd["Result"].ToString();
                    string TaskLocation = drusd["TaskLocation"].ToString();
                    string sign = drusd["Sign"].ToString();


                    string[] signtest = sign.Split(delimiterChars);
                    try
                    {
                        string signyear = signtest[1].ToString();
                        string signmonth = signtest[2].ToString();
                        string signdate = signtest[3].ToString();
                    }
                    catch { read = 1; }

                    string finish = drusd["FinishTime"].ToString();

                    string time = drusd["Time"].ToString();

                    if (read == 0 && finish.ToString().Contains(recordcombo.Text))
                    {
                        TaskCodeonly = drusd["TaskCardCode"].ToString();
                        timeend = drusd["FinishTime"].ToString();
                        sign_save = sign;
                        photodate = finish;
                        char[] delimiterChar1 = { ' ', ' ', '\r', '\a' };//remove \r\a
                        string[] photodatetest = photodate.Split(delimiterChar1);
                        photodate = photodatetest[0].ToString();
                        Recordgrid.Rows.Add(new Object[] { test, item, checkdetail, CheckResult, result, time, finish, sign, TaskLocation });
                    }

                }
                drusd.Close();
                cmdusd.Dispose();
                countrecord++;
            }
            
        }

        private void rejectbut_Click(object sender, EventArgs e)
        {
            if (!sign_save.ToString().Contains("reject"))
            {
                InputBox_reject input = new InputBox_reject();
                input.style = this.style;
                input.taskcode = Code.ToString();

                DialogResult dr = input.ShowDialog();
                if (dr == DialogResult.OK)
                {
                   
                    MySqlConnection myConnection = new MySqlConnection(connString);//連線資料庫

                    //開啟資料庫
                    myConnection.Open();
                    //2019 / 5 / 15 上午 11:51:09  讀出來格式
                    //2019-05-09 18:43:52

                    // MessageBox.Show(Code.ToString().ToLower());
                    string inser = "UPDATE `" + Code.ToString().ToLower() + "-f1-record" + @"` SET Sign=@changesign WHERE FinishTime = @sign ";

                    using (MySqlCommand cmd = new MySqlCommand(inser, myConnection))
                    {
                        Console.WriteLine(sign_save.ToString());
                        cmd.Parameters.AddWithValue("@sign", timeend.ToString());
                        cmd.Parameters.AddWithValue("@changesign", sign_save.ToString() + "/" + accountname.ToString() + "/reject");


                        int index = cmd.ExecuteNonQuery();
                        cmd.Dispose();




                    }
                    string sql = @"INSERT INTO `" + "taskrejected" + @"`(`TaskCard`,`TaskCardCode`,`Engineer`,`Manager`,`Reason`,`State`) VALUES
                                                       (@test1,@test2,@test3,@test4,@test5,@test6) ";

                    using (MySqlCommand cmd = new MySqlCommand(sql, myConnection))
                    {
                        cmd.Parameters.AddWithValue("@test3", Recordgrid.Rows[0].Cells[0].Value.ToString());
                        cmd.Parameters.AddWithValue("@test5", input.Getreson());

                        cmd.Parameters.AddWithValue("@test2", TaskCodeonly.ToString());
                        cmd.Parameters.AddWithValue("@test1", Code.ToString());
                        cmd.Parameters.AddWithValue("@test4", accountname.ToString());
                        cmd.Parameters.AddWithValue("@test6", "Rejected");



                        int index = cmd.ExecuteNonQuery();
                        cmd.Dispose();
                    }
                    myConnection.Close();
                }
                else if (dr == DialogResult.Cancel)
                {
                    //MessageBox.Show(input.GetchanMsg());
                }
                //關閉資料庫
            }
            else { MessageBox.Show("請勿重複退單"); }


        }

        private void signbut_Click(object sender, EventArgs e)
        {
       
            MySqlConnection myConnection = new MySqlConnection(connString);//連線資料庫

            //開啟資料庫
            myConnection.Open();
            //2019 / 5 / 15 上午 11:51:09  讀出來格式
            //2019-05-09 18:43:52

            // MessageBox.Show(Code.ToString().ToLower());
            string inser = "UPDATE `" + Code.ToString().ToLower() + "-f1-record" + @"` SET Sign=@changesign WHERE Sign= @sign ";

            using (MySqlCommand cmd = new MySqlCommand(inser, myConnection))
            {
                Console.WriteLine(sign_save.ToString());
                cmd.Parameters.AddWithValue("@sign", sign_save.ToString());
                cmd.Parameters.AddWithValue("@changesign", sign_save.ToString() + "/" + accountname.ToString() + "/"+ DateTime.Now.ToString("yyyy-MM-dd"));


                int index = cmd.ExecuteNonQuery();
                cmd.Dispose();




            }


            //關閉資料庫
            myConnection.Close();

        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
           
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (pictureBox1.Size.Width + 10 >= 680 && pictureBox1.Size.Height + 10 >= 630)
            {
                timer1.Enabled = false;
            }
            else
            {
                pictureBox1.Size = new Size(pictureBox1.Size.Width + 10, pictureBox1.Size.Height + 10);
                pictureBox1.Location = new Point(pictureBox1.Location.X + 7, 30);
            }
            // MessageBox.Show(sizex.ToString() + sizey.ToString());


        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (photocount == 0)
            { }
            else
            {//多張圖片展示 1
                if (photousecount < photocount)
                {
                    Console.WriteLine(photoname[photousecount]);
                    pictureBox1.Image = Image.FromFile(photoname[photousecount]);
                    timer1.Enabled = true;
                    photousecount++;
                }
               else if (photousecount == photocount)
                { photousecount = 0; }
                
            }


        }

        private void outbut_Click(object sender, EventArgs e)//匯出表單
        { //Graphics g = pictureBox5.CreateGraphics();
            Graphics g = Graphics.FromImage(this.pictureBox5.Image);
            // drawing = true;
            g.DrawString("thsrc", Font, Brushes.Red, 25, 10);
            var taiwanCalendar = new System.Globalization.TaiwanCalendar();
            var datatime3 = DateTime.Now.ToString("MM.dd.yyyy");
            var datatime1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            
            g.DrawString(taiwanCalendar.GetYear(datatime1).ToString() + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day, this.Font, Brushes.Red, 15, 35);
            string engin = Recordgrid.Rows[0].Cells[0].Value.ToString();
            g.DrawString(engin.ToString(), this.Font, Brushes.Red, 15, 60);
        
            pictureBox5.Image.Save(@"D:\SCD\Sign\"+ engin.ToString() + taiwanCalendar.GetYear(datatime1).ToString() + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".jpg", ImageFormat.Jpeg);
            pictureBox5.Image = pictureBox5.Image;
            string thisdaypath = @"D:\SCD\Sign\" + engin.ToString() + taiwanCalendar.GetYear(datatime1).ToString() + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".jpg";
            //主管簽
            //Graphics g1 = Graphics.FromImage(this.pictureBox5.Image);
            // drawing = true;
          
                Graphics g1 = Graphics.FromImage(this.pictureBox6.Image);
                g1.DrawString("thsrc", Font, Brushes.Red, 25, 10);
                g1.DrawString(taiwanCalendar.GetYear(datatime1).ToString() + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day, this.Font, Brushes.Red, 15, 35);
                string engin1 = Recordgrid.Rows[0].Cells[7].Value.ToString();
                string[] enginx = engin1.Split('/');
                Console.WriteLine(enginx[0].ToString());
                g1.DrawString(enginx[0].ToString(), this.Font, Brushes.Red, 15, 60);

                pictureBox6.Image.Save(@"D:\SCD\Sign\" + enginx[0].ToString() + taiwanCalendar.GetYear(datatime1).ToString() + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".jpg", ImageFormat.Jpeg);
                pictureBox6.Image = pictureBox5.Image;
                string thisdaypath1 = @"D:\SCD\Sign\" + enginx[0].ToString() + taiwanCalendar.GetYear(datatime1).ToString() + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + ".jpg";


            //匯出
            string Pathcatch = "";//檔案搜尋路徑位置
            string TaskName = "";
            //string  = "";
            //string standard = "";
            string note = "";
         
            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

            //有表格
            string sql = @"select * from `" + Code.ToString().ToLower() + "`";

            MySqlCommand cmdusd = new MySqlCommand(sql, conn);
            MySqlDataReader drusd = cmdusd.ExecuteReader();
            while (drusd.Read())
            {
                TaskName = drusd["TaskCardName"].ToString();
            }
            try
            {
                foreach (Process item in Process.GetProcessesByName("WINWORD"))
                { item.Kill(); }//如存在開啟的Word則先關閉(一個應用程式例項就是程序)
//注意工卡位置
               
                foreach (string fname in System.IO.Directory.GetFileSystemEntries(@"D:\SCD\下載\AllTask\"))
                {
                    if (fname.ToString().Contains(TaskName.ToString()))
                    {//抓路徑
                        Pathcatch = Path.GetFullPath(fname.ToString());
                        Filename = Path.GetFileNameWithoutExtension(fname.ToString());
                        Console.WriteLine(Pathcatch);
                        break;
                    }
                }
                //圖片載
                
               /* Document doc = new Document();
                doc.LoadFromFile(Pathcatch.ToString());
                Table table1 = (Table)doc.Sections[1].Tables[0];
                table1.Rows[1].Cells[5].Paragraphs[0].Text = "";
                DocPicture picture = table1.Rows[1].Cells[5].Paragraphs[0].AppendPicture(Image.FromFile(thisdaypath.ToString()));
                
                doc.SaveToFile(Pathcatch.ToString(), FileFormat.Docx);*/
foreach (Process item in Process.GetProcessesByName("WINWORD"))
                { item.Kill(); }//如存在開啟的Word則先關閉(一個應用程式例項就是程序)
//資料匯入

                object filePath = Pathcatch;
                wordApp = new Word.Application(); //應用程式例項化
                                                  //wordApp.Visible = true;//可見(用的是wordApp顯示)
                wordDoc = wordApp.Documents.Open(ref filePath, ref Nothing, ref IsReadOnly, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);

                wordDoc.Tables[1].Select();//選擇與複製：文件用Sections[節]/應用程式用Selection[選擇]
                wordApp.Selection.Copy();
               
                object myType = Word.WdBreakType.wdSectionBreakContinuous;//換行符
                object myUnit = Word.WdUnits.wdStory;
                object pBreak = (int)Word.WdBreakType.wdPageBreak;//下一頁(上一頁最後一行空,否則跳入下下頁)

               
                for (int i = 0; i < tableCount - 1; i++)//原本手簿已存在一個，則複製減1個
                {
                    wordApp.Selection.EndKey(ref myUnit, ref Nothing);
                    wordApp.Selection.InsertBreak(ref pBreak);
                    wordApp.Selection.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault);//複製語句：應用程式形式
                                                                                         //wordApp.Selection.Paste();//複製語句：簡單形式
                }
             Console.WriteLine("手簿載入完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            drusd.Dispose();
            cmdusd.Dispose();
            conn.Close();
            try
            {
                int tabturn = 0;//表格計數
                foreach (Word.Table table in wordDoc.Tables)
                {
                    #region 文件表格賦值
                    if (tabturn == 2)
                    {
                    }

                    if (tabturn == 3)//紀錄的表格位置
                    { int errorcount = 1;
                        int unerrorcount = 1;
                      for (int i = 0; i < Recordgrid.RowCount - 1; ++i)
                            {
                                
                                table.Cell(i + 4, 4).Range.Text = Recordgrid.Rows[i].Cells[3].Value.ToString();

                            if (Recordgrid.Rows[i].Cells[4].Value.ToString().Equals("Qualified"))
                            {
                                try//正常的狀態下
                                {    table.Cell(i + 4, 6).Range.Text = "■";
                                    try { table.Cell( i+4-errorcount, 6).Range.Text = "■"; }//都正常財部會到下一行
                                    catch { table.Cell(i + 4 - errorcount-unerrorcount+1, 7).Range.Text = "■"; errorcount = 1;unerrorcount = 1; }
                                }
                                catch
                                {
                                    errorcount++;
                                }
                            }
                            else
                                try//不正常的狀態下
                                {
                                    table.Cell(i + 4, 7).Range.Text = "■";
                                    try { table.Cell(i + 4 - unerrorcount, 7).Range.Text = "■"; }//都不正常才部會到下一行
                                    catch { table.Cell(i + 4 - errorcount - unerrorcount + 1, 6).Range.Text = "■"; errorcount = 1; unerrorcount = 1; }
                                }
                                catch
                                {
                                    unerrorcount++;
                                }
                            
                                //   MessageBox.Show(i.ToString() + "\n" + Recordgrid.Rows[i].Cells[4].Value.ToString());
                            }
                                           
                    }

                    #endregion

                    tabturn += 1;
                }

                //文件另存為並退出
                //object filePath = @"D:\SCD\TaskRecord記錄位置\test.docx";
                Console.WriteLine(recordcombo.Text.ToString());
                string Date = recordcombo.Text.ToString().Replace("-", "_").Substring(0,10);
                Console.WriteLine(Date);
                object filePath = @"D:\SCD\TaskRecord記錄位置\"+ Filename.ToString()+Date.ToString()+".docx";
                //  object format = Word.WdSaveFormat.wdFormatDocumentDefault;//docx
                wordDoc.SaveAs(ref filePath, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
                wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);                                                         //object format = MSWord.WdSaveFormat.wdFormatDocument;//doc


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (wordApp != null)
                {
                    Marshal.FinalReleaseComObject(wordApp);
                    wordApp = null;
                }

                if (wordDoc != null)
                {
                    Marshal.FinalReleaseComObject(wordDoc);
                    wordDoc = null;
                }
                foreach (Process item in Process.GetProcessesByName("WINWORD"))
                { item.Kill(); }//如存在開啟的Word則先關閉(一個應用程式例項就是程序)
                                //注意工卡位置

                foreach (string fname in System.IO.Directory.GetFileSystemEntries(@"D:\SCD\TaskRecord記錄位置\"))
                {
                    if (fname.ToString().Contains(TaskName.ToString()) && fname.ToString().Contains(recordcombo.Text.ToString().Replace("-", "_").Substring(0, 10)))
                    {//抓路徑
                        Pathcatch = Path.GetFullPath(fname.ToString());
                        Filename = Path.GetFileNameWithoutExtension(fname.ToString());
                        Console.WriteLine(Pathcatch);
                        break;
                    }
                }
                //圖片載

                 Document doc = new Document();
                 doc.LoadFromFile(Pathcatch.ToString());
                 Table table1 = (Table)doc.Sections[1].Tables[0];
                 table1.Rows[1].Cells[5].Paragraphs[0].Text = "";
                table1.Rows[1].Cells[4].Paragraphs[0].Text = "";
                table1.Rows[1].Cells[3].Paragraphs[0].Text = "";
                DocPicture picture = table1.Rows[1].Cells[5].Paragraphs[0].AppendPicture(Image.FromFile(thisdaypath.ToString()));
                DocPicture picture1 = table1.Rows[1].Cells[4].Paragraphs[0].AppendPicture(Image.FromFile(thisdaypath.ToString()));
                DocPicture picture2 = table1.Rows[1].Cells[3].Paragraphs[0].AppendPicture(Image.FromFile(thisdaypath1.ToString()));
                doc.SaveToFile(Pathcatch.ToString(), FileFormat.Docx);
            }



        }
    }
}
