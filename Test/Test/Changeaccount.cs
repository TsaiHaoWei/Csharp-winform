using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
    public partial class Changeaccount : MetroFramework.Forms.MetroForm
    {
        public int style = 0;
        public string Name = "";
        public string IP = "";
        public string port = "";
        public Changeaccount()
        {
			
            InitializeComponent();
           
        }

        private void Changeaccount_Load(object sender, EventArgs e)
        {
            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);
             changeuser.Text= Name;
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            if (oldsecret.Text != String.Empty && newsecret.Text != String.Empty)
            {
                string connString = "server=" + IP.ToString() + ";port=" + port.ToString() + ";user id=thsr2019_05_08;password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫s
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
 //抓密碼
                String select = @"select Password from `account` where Name = @user";
                    MySqlCommand cmd = new MySqlCommand(select, conn);
                    cmd.Parameters.Add(new MySqlParameter("@user", Name.ToString()));
                    MySqlDataReader dr = cmd.ExecuteReader();

                    if (dr.HasRows) //如果有抓到資料
                    {
                        SHA256 sha256 = new SHA256CryptoServiceProvider();//建立一個SHA256
                        byte[] source = Encoding.Default.GetBytes(oldsecret.Text);//將字串轉為Byte[]
                        byte[] crypto = sha256.ComputeHash(source);//進行SHA256加密
                        string result = Convert.ToBase64String(crypto);//把加密後的字串從Byte[]轉為字串

                    try {
                        while ((dr.Read()))

                        {    //5.判斷資料列是否為空

                            if (!dr[0].Equals(DBNull.Value))

                            {
                                if (dr[0].ToString().Equals(result))
                                {///密碼正確

                                    dr.Close();
                                    cmd.Dispose();
                                    string up = "UPDATE `account` SET Password = @p1 WHERE Name =  @n1 ";
                                    using (MySqlCommand update = new MySqlCommand(up, conn))
                                    {
                                        // 加密
                                        SHA256 sha = new SHA256CryptoServiceProvider();//建立一個SHA256
                                        byte[] chsource = Encoding.Default.GetBytes(newsecret.Text);//將字串轉為Byte[]
                                        byte[] chcrypto = sha.ComputeHash(chsource);//進行SHA256加密
                                        string chresult = Convert.ToBase64String(chcrypto);//把加密後的字串從Byte[]轉為字串
                                        Console.WriteLine(chresult.ToString());

                                        update.Parameters.Add("@p1", chresult.ToString());
                                        update.Parameters.Add("@n1", Name);
                                        int index = update.ExecuteNonQuery();
                                        if (index > 0)
                                        {
                                            MetroFramework.MetroMessageBox.Show(this, "密碼更改成功", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            this.Visible = false;
                                            update.Dispose();
                                        }
                                    }


                                }
                                else
                                    MessageBox.Show("密碼錯誤");
                                oldsecret.Text = "";
                                newsecret.Text = "";
                            }

                        }

                    }

                   catch { }


                }


                    else
                        MessageBox.Show("查無此帳號");


                    dr.Dispose();
                    cmd.Dispose();
                    conn.Close();
                    conn.Dispose();

                }
              
            
            else MetroFramework.MetroMessageBox.Show(this, "請先輸入密碼", "錯誤輸入", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }

      
    }
}
