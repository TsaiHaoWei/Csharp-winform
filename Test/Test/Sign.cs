using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
    public partial class Sign : MetroFramework.Forms.MetroForm
    {
        public string IP = "";//預設的資料庫IP port
        public string port = "";

        public int style = 0;

        string position = "";
       
        //string[] fileName = new string[5];
        string filename = "";
        string connString = "";
        public Sign(int style)
        {       InitializeComponent();

            this.style = style;
            this.StyleManager = metroStyleManager1;
          
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);
        }

        private void Sign_Load(object sender, EventArgs e)
        {
           connString = "server=" + IP.ToString() + ";port=" + port.ToString() + ";user id=thsr2019_05_08;password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫
            Console.WriteLine(connString.ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtName.Text.ToString() != String.Empty && txtPassword.Text.ToString() != String.Empty
                && comboBox1.Text.ToString() != String.Empty)
            {
                
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

                String select = @"select Name from `account` where Name = @user";
                // DataTable dt = new DataTable();
                //              MySqlDataAdapter adapter = new MySqlDataAdapter(select, conn);
                MySqlCommand cmd1 = new MySqlCommand(select, conn);
                cmd1.Parameters.Add(new MySqlParameter("@user", txtName.Text));
                MySqlDataReader dr = cmd1.ExecuteReader();
                if (dr.HasRows) MessageBox.Show("此帳號有人使用");//有抓到資料

                else
                {
                    dr.Close();
                    string sql = @"INSERT INTO `account` (`Name`, `Password`, `Team`,`Position`) VALUES
                              (@test1,@test2,@test4,@test5) ";
                    MySqlCommand cmd = new MySqlCommand(sql, conn);
                    // 加密
                    SHA256 sha256 = new SHA256CryptoServiceProvider();//建立一個SHA256
                    byte[] source = Encoding.Default.GetBytes(txtPassword.Text);//將字串轉為Byte[]
                    byte[] crypto = sha256.ComputeHash(source);//進行SHA256加密
                    string result = Convert.ToBase64String(crypto);//把加密後的字串從Byte[]轉為字串

                    cmd.Parameters.Add("@test1", txtName.Text);
                    cmd.Parameters.Add("@test2", result.ToString());
                    cmd.Parameters.Add("@test4", comboBox1.Text.ToString());
                    cmd.Parameters.Add("@test5", position.ToString());
                   


                    int index = cmd.ExecuteNonQuery();
                    bool success = false;
                    if (index > 0)
                    {    //   MessageBox.Show("註冊成功");
                       
                   

                        string rolelogin = "CREATE USER '"+ txtName.Text + "'@'%' IDENTIFIED BY 'CzqhTlz0erd13UX6';";
                      
                                                                                                                                                                                                                                                                                                                                                                                                                                         //     string sql = "sp_adduser 'ulong', 'ulong', 'db_owner';


                        using (MySqlConnection myConn = new MySqlConnection(connString))
                        {

                            MySqlCommand myCommand = new MySqlCommand(rolelogin, myConn);
                            try
                            {
                                myConn.Open();
                                myCommand.ExecuteNonQuery();
                                MessageBox.Show("創建成功", "MyProgram", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (System.Exception ex)
                            {
                                MessageBox.Show(ex.ToString(), "MyProgram", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            finally
                            {
                                if (myConn.State == ConnectionState.Open)
                                {
                                    myConn.Close();
                                }
                            }
                        }
                        Form1 f1 = new Form1();
                        f1.IP = this.IP;
                        f1.port = this.port;
                        f1.style = style;
                        this.Visible = false;
                        f1.Visible = true;
                    }
                    else
                        MessageBox.Show("註冊失敗");
                }
            }
            else MessageBox.Show("請勿空白");
            txtName.Text = ""; txtPassword.Text = ""; comboBox1.Text = "";
        }

      
        private void button2_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.IP = this.IP;
            f1.port = this.port;
            f1.style = style;
               this.Visible = false;
            f1.Visible = true;
        }

         private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            position = "backmanger";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            position = "backmanager";
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            position = "manager";
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            position = "manager";
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            position = "manager";
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            position = "engineer";
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            position = "engineer";
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog oOpenFileDialog = new OpenFileDialog())
            {
                oOpenFileDialog.Filter = " image|*.JPG | image|*.JPEG | image|*.PNG| All Files|*.*";
                oOpenFileDialog.Title = "Select a  File";
                oOpenFileDialog.FilterIndex = 3;


                if (oOpenFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                  string  path = oOpenFileDialog.FileName;//絕對路徑

                    filename = Path.GetFileName(path);
                    Console.WriteLine(filename);
                    // fileName[count] = Path.GetFileName(path);//設定路徑跟檔名 資料庫用
                  //  Filepath[count] = "ftp://163.18.57.239/" + fileName[count];
                    pictureBox1.Image = Image.FromFile(path);
                    // pictureBox1.BackColor = System.Drawing.Color.Transparent;
                    string photoname = txtName.Text +".jpg";

                    string fs = @"D:\SCD\Sign" + photoname.ToString();
                    pictureBox1.Image.Save(fs, System.Drawing.Imaging.ImageFormat.Png);
                    FTPUpload(photoname.ToString());


                }

            }

        }
        private void FTPUpload(string photoname)
        {

            string username = "test1";
            string password = "1234";
            string uploadUrl = "ftp://163.18.57.239/" + "Sign" + "/" + photoname;

            if (!System.IO.File.Exists(@"D:\SCD\Sign" + photoname.ToString()))
                MessageBox.Show("上傳檔案不存在");

            //   MessageBox.Show(i.ToString());
            // 要上傳的檔案
            //圖片格式
            string MyFileName = @"D:\SCD\Sign" + photoname.ToString();
            //文件格式擋
            /*StreamReader sourceStream = new StreamReader(@"C:\Users\B510\Desktop\File\" + i.ToString()+".jpg");
            byte[] data = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
            sourceStream.Close();
            */
            WebClient wc = new WebClient();
            wc.Credentials = new NetworkCredential(username, password);
            wc.UploadFile(uploadUrl, MyFileName);
        //    path = uploadUrl.ToString();
            ///文件格式上傳 wc.UploadData(uploadUrl, data);

        }
    }
}
