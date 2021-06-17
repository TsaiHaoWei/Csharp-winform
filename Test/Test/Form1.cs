using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using MySql.Data.MySqlClient;
using System.Security.Cryptography;
using System.Drawing.Drawing2D;
using System.Net;
using System.IO;
using System.Drawing.Imaging;
using System.Diagnostics;
//引用API
using System.Runtime.InteropServices;
using System.Data.SqlClient;

namespace Test
{
   
    

    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
       [DllImport("kernel32")]
    private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
    [DllImport("kernel32")]
    private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        public CookieContainer cc = new CookieContainer();//連線WEBSERVE
        //AutoSizeFormClass asc = new AutoSizeFormClass();    
        public string IP = "";//預設的資料庫IP port
        public string port = "";
        //FTP
        string ftpServerIP = "";
        string ftpUserID = "";
        string ftpPassword = "";
        FtpWebRequest reqFTP;

        public int style = 0;

        public string name = "";
        string team = "";
        string permission = "";

        
        public string[] taskname = new string[100];

        //匯入taskcard資料用的
        string Good;
        string Item;
        string Detail;
        string Detail1 = "";//項目細節輸出
        string standard;
        string Note;
        // int i = 0;//進去資料庫的圖檔名
        string photoname = "";
        string path = "";
        string fileName = "";//檔案名
        int index = 0;//判斷資料庫有沒有進去
        int count = 1;//圖片匯入的順序
        string[] words = { };//更改後的程式
        string lower = "";
        string nametask = "";



      

        string connString = "";
        string allconnString = "";

        
        public Form1()////Connect_problem
        {
            //Read ini File


            StringBuilder retVal = new StringBuilder(255);  //回傳所要接收的值


            string Section = "SQL serve";
            string[] Key = { "IP", "port", "FTP IP", "FTP User", "FTP Password", "style" };
            string Defaut = "null";      //如果沒有 Section , Key 兩個參數值，則將此值賦給變量
            int Size = 255;              //設定回傳 Siez 

            
            //Console.WriteLine(Open.FileName);
             for (int i = 0; i < Key.Length; i++)
              { switch (i)
                   {
                      case 0:
                            int  strref = GetPrivateProfileString(Section, Key[i], Defaut, retVal, Size, @"D:\SCD\Record.ini");
                            IP = retVal.ToString();
                             Console.WriteLine(IP.ToString());
                            break;
                      case 1:
                            strref = GetPrivateProfileString(Section, Key[i], Defaut, retVal, Size, @"D:\SCD\Record.ini");
                            port = retVal.ToString();
                            break;
                      case 2:
                            strref = GetPrivateProfileString(Section, Key[i], Defaut, retVal, Size, @"D:\SCD\Record.ini");
                             ftpServerIP = retVal.ToString();
                            break;
                      case 3:
                            strref = GetPrivateProfileString(Section, Key[i], Defaut, retVal, Size, @"D:\SCD\Record.ini");
                             ftpUserID = retVal.ToString();
                             break;
                    case 4:
                            strref = GetPrivateProfileString(Section, Key[i], Defaut, retVal, Size, @"D:\SCD\Record.ini");
                            ftpPassword = retVal.ToString();
                            break;
                    case 5:
                            strref = GetPrivateProfileString(Section, Key[i], Defaut, retVal, Size, @"D:\SCD\Record.ini");
                            style = Int32.Parse(retVal.ToString());
                            break;
                    }

              }


       

            

            allconnString = "server=" + IP.ToString() + ";port=" + port.ToString() + ";user id=thsr2019_05_08;password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫
            InitializeComponent();
           
        }
        private void Form1_Load(object sender, EventArgs e)//Style  版面大小 
        {//版面大小 
            this.Size = new System.Drawing.Size(Screen.PrimaryScreen.WorkingArea.Width - 100, Screen.PrimaryScreen.WorkingArea.Height - 100);
            ///Style
            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);

        }
        private void button1_Click(object sender, EventArgs e)//登入////Connect_problem
        {
            if (metroTextBox1.Text != String.Empty)
            {//抓權限
                connString = "server=" + IP.ToString() + ";port=" + port.ToString() + ";user id=" + metroTextBox1.Text + ";password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫 
               
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
                try
                {
                    if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                    String selectused = @"select Permission from `account` where Name = @user";
                    MySqlCommand cmdusd = new MySqlCommand(selectused, conn);
                    cmdusd.Parameters.Add(new MySqlParameter("@user", metroTextBox1.Text));
                    MySqlDataReader drusd = cmdusd.ExecuteReader();
                    while (drusd.Read())
                    {
                        if (!drusd[0].Equals(DBNull.Value))
                        {
                            permission = drusd[0].ToString();
                        }
                    }
                    Console.WriteLine(permission);
                    drusd.Close();
                    cmdusd.Dispose();

                    if (permission.ToString() != String.Empty)
                    {//抓組別
                        String selectGroup = @"select Team from `account` where Name = @user";
                        MySqlCommand cmd1 = new MySqlCommand(selectGroup, conn);
                        cmd1.Parameters.Add(new MySqlParameter("@user", metroTextBox1.Text));
                        MySqlDataReader dr1 = cmd1.ExecuteReader();
                        while (dr1.Read())
                        {
                            if (!dr1[0].Equals(DBNull.Value))
                            {
                                team = dr1[0].ToString();
                            }
                        }

                        dr1.Close();
                        cmd1.Dispose();
                        String select = @"select Password from `account` where Name = @user";
                        MySqlCommand cmd = new MySqlCommand(select, conn);
                        cmd.Parameters.Add(new MySqlParameter("@user", metroTextBox1.Text));
                        MySqlDataReader dr = cmd.ExecuteReader();

                        if (dr.HasRows) //如果有抓到資料
                        {
                            SHA256 sha256 = new SHA256CryptoServiceProvider();//建立一個SHA256
                            byte[] source = Encoding.Default.GetBytes(metroTextBox2.Text);//將字串轉為Byte[]
                            byte[] crypto = sha256.ComputeHash(source);//進行SHA256加密
                            string result = Convert.ToBase64String(crypto);//把加密後的字串從Byte[]轉為字串
                            while ((dr.Read()))

                            {    //5.判斷資料列是否為空

                                if (!dr[0].Equals(DBNull.Value))

                                {
                                    if (dr[0].ToString().Equals(result))
                                    {
                                        Form2 f = new Form2(metroTextBox1.Text, team, permission, style);//產生Form2的物件，才可以使用它所提供的Method
                                                                                                         //f.StyleManager = this.StyleManager;
                                        f.IP = this.IP;
                                        f.port = this.port;//f.ShowDialog();
                                                           // f.Dispose();
                                        this.Visible = false;
                                        f.ShowDialog();
                                        metroTextBox1.Text = "";
                                        metroTextBox2.Text = "";

                                    }
                                    else
                                        MessageBox.Show("密碼錯誤");
                                    metroTextBox2.Text = "";
                                }

                            }

                        }


                        else
                            MessageBox.Show("查無此帳號");


                        dr.Dispose();
                        cmd.Dispose();
                        conn.Close();
                        conn.Dispose();

                    }
                    else
                        MessageBox.Show("未有登入權限", "錯誤警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch
                {
                    MessageBox.Show("未有登入權限", "錯誤警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                
               
            }
            else MessageBox.Show("請先輸入帳號密碼", "錯誤警告", MessageBoxButtons.OK, MessageBoxIcon.Error);


        }

        private void button2_Click(object sender, EventArgs e)//註冊
        {
         
            Sign f = new Sign(style);//產生Form2的物件，才可以使用它所提供的Method
            f.IP = this.IP;
            f.port = this.port;
           
            this.Visible = false;
            f.Visible = true; 

               
        }
        
        private void pictureBox4_Click(object sender, EventArgs e)//關閉城市
        {
            System.Environment.Exit(0);
        }
        
        private void pictureBox3_Click(object sender, EventArgs e)//設置顏色
        {
            metroPanel1.Visible = true;
        }


        private void button4_Click(object sender, EventArgs e)//red
        {
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(13);
            style = 13;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "13";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
            
        }


        private void button7_Click(object sender, EventArgs e)//blue+green
        {
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Teal;
            style = 7;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "7";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
        }

        private void button8_Click(object sender, EventArgs e)//blue
        {
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Blue;
            style = 4;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "4";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
        }

        private void button9_Click(object sender, EventArgs e)//GREEN
        {
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Green;
            style = 5;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "5";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
        }

        private void button10_Click(object sender, EventArgs e)//LIGHT +green
        {
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Lime;
            style = 6;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "6";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
        }

        private void button15_Click(object sender, EventArgs e)//orange
        {
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Orange;
            style = 8;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "8";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
        }

        private void button11_Click(object sender, EventArgs e)//PINK
        {
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Pink;
            style = 10;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "10";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
        }

        private void button12_Click(object sender, EventArgs e)//purple
        {
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Purple;
            style = 12;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "12";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
        }

        private void button13_Click(object sender, EventArgs e)//YELLO
        {
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Yellow;
            style = 14;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "14";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
        }

        private void button14_Click(object sender, EventArgs e)//LIGHT purple
        {
            metroStyleManager1.Style = MetroFramework.MetroColorStyle.Magenta;
            style = 11;
            //Write ini File
            string Section = "SQL serve";
            string Key = "style";
            string Value = "11";

            WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");
        }
        private void button3_Click_1(object sender, EventArgs e)//更改IP////Connect_problem
        {
            try
            {
                string testconnString = "Data Source=163.18.57.239;Database=thsr v1;User ID=Ta;Password=1234";//連線資料庫
            
                /*string testconnString = "server=" + IPtxt.Text + ";port=" + portxt.Text + ";user id=thsr2019_05_08;password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫
             */
                using (MySqlConnection conn = new MySqlConnection(testconnString))//實做一個物件
                {
                   
                    //成功修改
                    IP = IPtxt.Text;
                    port = portxt.Text;
                    //修改ini
                    string Section = "SQL serve";
                    string Key = "IP";
                    string Value = IP;
                     WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");

                    Section = "SQL serve";
                    Key = "port";
                    Value = port;
                    WritePrivateProfileString(Section, Key, Value, @"D:\SCD\Record.ini");

                    MessageBox.Show("更改成功");
                    allconnString = testconnString.ToString();
                    IPtxt.Text = "";
                    portxt.Text = "";
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                MetroFramework.MetroMessageBox.Show(this, "尚未連接成功", "error");
                IPtxt.Text = "";
                portxt.Text = "";
            }




        }

        private void butPic_Click(object sender, EventArgs e)//縮小調整畫面
        {
            metroPanel1.Visible = false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            labTime.Text = DateTime.Now.ToString();

           
        }
      

        private void portxt_Click(object sender, EventArgs e)//快速紐
        {
            this.AcceptButton = button3;
        }

        private void metroTextBox2_Click(object sender, EventArgs e)//快速紐
        {
            this.AcceptButton = button1;
        }

        private void metroTextBox1_Click(object sender, EventArgs e)//快速紐
        {
            this.AcceptButton = button1;

        }

        private void portxt_Leave(object sender, EventArgs e)//快速紐
        {
            this.AcceptButton = button3;

        }

        private void IPtxt_Leave(object sender, EventArgs e)//快速紐
        {
            this.AcceptButton = button3;

        }
  


       

        private void button6_Click_1(object sender, EventArgs e)//刪除本基WORD資料
        {
        
       
            foreach (string fname in System.IO.Directory.GetFileSystemEntries(@"D:\SCD\下載\TaskDownload\"))
            {
                Console.WriteLine(fname.ToString());
            }
            foreach (Process item in Process.GetProcessesByName("WINWORD"))
            { item.Kill(); }//如存在開啟的Word則先關閉(一個應用程式例項就是程序)
                            //刪除本機端電腦檔案
            string file = @"D:\SCD\下載\TaskDownload\";
            //去除資料夾和子檔案的只讀屬性
            //去除資料夾的只讀屬性
            System.IO.DirectoryInfo fileInfo = new DirectoryInfo(file);
            fileInfo.Attributes = FileAttributes.Normal & FileAttributes.Directory;
            //去除檔案的只讀屬性
            System.IO.File.SetAttributes(file, System.IO.FileAttributes.Normal);
            //判斷資料夾是否還存在
            if (Directory.Exists(file))
            {
                foreach (string f in Directory.GetFileSystemEntries(file))
                {
                    if (File.Exists(f))
                    {
                        //如果有子檔案刪除檔案
                        File.Delete(f);
                    }
                    /*lse
                     {
                         //迴圈遞迴刪除子資料夾 
                         DeleteSrcFolder1(f);
                     }*/
                }
                //刪除空資料夾
                //Directory.Delete(file);
            }
            //刪除FTP 上的檔案

        }

        private void button5_Click(object sender, EventArgs e)//連線WEBSERVE
        {
            string url = "http://163.18.57.236/login.php";

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

            request.Method = "POST";

            request.ContentType = "application/x-www-form-urlencoded";

            request.CookieContainer = cc;

            string data;



            string user = "Tony"; //用戶名

            string pass = "3WVw1Fz8Jdfgvsi07xHaSI2YPQpfp7MPgF7AHwXorrY="; //密碼

            data = "&UserName=" + user + "&Password=" + pass;

            request.ContentLength = data.Length;

            StreamWriter writer = new StreamWriter(request.GetRequestStream(), Encoding.ASCII);

            writer.Write(data);

            writer.Flush();





            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            string encoding = response.ContentEncoding;

            if (encoding == null || encoding.Length < 1)

            {

                encoding = "UTF-8"; //預設編碼

            }

            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding(encoding));

            data = reader.ReadToEnd();

            MessageBox.Show(data);

            cc = request.CookieContainer;

            response.Close();
        }
    }
}
