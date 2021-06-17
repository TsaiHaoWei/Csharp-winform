using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
    public partial class Lock : MetroFramework.Forms.MetroForm
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        public string Password = "";
        public string Name = "";


        //  
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
        public Lock()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
   
                SHA256 sha256 = new SHA256CryptoServiceProvider();//建立一個SHA256
                byte[] source = Encoding.Default.GetBytes(textBox1.Text);//將字串轉為Byte[]
                byte[] crypto = sha256.ComputeHash(source);//進行SHA256加密
                string result = Convert.ToBase64String(crypto);//把加密後的字串從Byte[]轉為字串
            if (Password.ToString().Equals(result))
            {
                textBox1.Text = "";
                this.Visible = false;
            }
            else
            { textBox1.Text = "";
                MetroFramework.MetroMessageBox.Show(this, "密碼錯誤");
         }
             
        }

        private void Lock_Load(object sender, EventArgs e)
        {
            StringBuilder retVal = new StringBuilder(255);  //回傳所要接收的值


            string Section = "SQL serve";
            string[] Key = { "IP", "port", "FTP IP", "FTP User", "FTP Password", "style" };
            string Defaut = "null";      //如果沒有 Section , Key 兩個參數值，則將此值賦給變量
            int Size = 255;              //設定回傳 Siez 


            //Console.WriteLine(Open.FileName);
            for (int i = 0; i < Key.Length; i++)
            {
                switch (i)
                {
                    case 0:
                        int strref = GetPrivateProfileString(Section, Key[i], Defaut, retVal, Size, @"D:\SCD\Record.ini");
                        IP = retVal.ToString();
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


        }
        public void DeleteFileName(string fileName)//刪除FTP 上的WORD檔案

        {
            try

            {
                FileInfo fileInf = new FileInfo(fileName);

                string uri =  ftpServerIP + "MIS輸出檔/" + fileInf;

                Connect(uri);//连接        

                // 默认为true，连接不会被关闭

                // 在一个命令之后被执行

                reqFTP.KeepAlive = false;

                // 指定执行什么命令

                reqFTP.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                response.Close();

            }

            catch (Exception ex)

            {

                  //MessageBox.Show(ex.Message, "删除错误");

            }

        }
        public void Taskcard()
        {

            Microsoft.Office.Interop.Word._Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word._Document doc = null;

            object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument;
            object unknow = Type.Missing;
            // app.Visible = true;
            object oFileName = @"" + path + "";


            doc = app.Documents.Open(ref oFileName, ref format,//資料路徑 格式
                       ref unknow, ref unknow, ref unknow, ref unknow,
                       ref unknow, ref unknow, ref unknow, ref unknow,
                       ref unknow, ref unknow, ref unknow, ref unknow,
                       ref unknow, ref unknow);

            int count = 0;
            ///抓WORD表格
            for (int tablePos = 2; tablePos <= 2; tablePos++)//抓第幾個表格
            {
                Microsoft.Office.Interop.Word.Table nowTable = doc.Tables[1];
                int deviceconstant = 0;
                string Alltaskcode = "";
                string Alltaskname = "";
                string Alltaskdetail = "";
                string Alltasklocation = "";
                string Alltaskdevice = "";
                string maintain = "";
                string safe = "";
                //需求
                string people = "";
                string material = "";
                string Device = "";
                //爭測哪種需求
                int need = 0;
                int rowsql = 0;
                for (int rowPos = 1; rowPos <= nowTable.Rows.Count; rowPos++)//抓第幾----
                {


                    for (int columPos = 1; columPos <= nowTable.Columns.Count; columPos++)//抓底幾個|||
                    {

                        //需求設定
                        try//因為表格每一行欄位數量不一致
                        {
                            Good = nowTable.Cell(rowPos, columPos).Range.Text;
                            Good = Good.Remove(Good.Length - 2, 2);

                            switch (Good.ToString())
                            {
                                case "人員需求":
                                    need = 1;
                                    rowsql = rowPos + 2;
                                    break;
                                case "人員需求Manning Requirement":
                                    need = 1;
                                    rowsql = rowPos + 2;
                                    break;
                                case "材料需求":
                                    need = 2;
                                    rowsql = rowPos + 2;
                                    break;
                                case "材料需求Materials Requirement":
                                    need = 2;
                                    rowsql = rowPos + 2;
                                    break;
                                case "工具/治具/機具設備需求":
                                    need = 3;
                                    rowsql = rowPos + 2;
                                    deviceconstant++;
                                    break;

                                case "工具/治具/機具設備需求 Tool/Jig/Equipment Requirement":
                                    need = 3;
                                    rowsql = rowPos + 2;
                                    deviceconstant++;
                                    break;
                                default:
                                    break;
                            }
                            if (need == 1 & rowsql <= rowPos)
                            {
                                if (!(Good.ToString().Equals("") || Good.ToString().Contains("□")))
                                {
                                    switch (columPos)
                                    {
                                        case 1:
                                            people += "~";
                                            people += Good.ToString();
                                            break;
                                        case 2:
                                            people += "!";
                                            people += Good.ToString();
                                            break;
                                        case 3:
                                            people += "@";
                                            people += Good.ToString();
                                            break;
                                        case 4:
                                            people += "#";
                                            people += Good.ToString();
                                            break;
                                        case 5:
                                            people += "$";
                                            people += Good.ToString();
                                            break;
                                        case 6:
                                            people += "%";
                                            people += Good.ToString();
                                            break;
                                        case 7:
                                            people += "^";
                                            people += Good.ToString();
                                            break;
                                        default:
                                            MessageBox.Show("default");
                                            break;

                                    }

                                }
                            }



                            else if (need == 2 & rowsql <= rowPos)
                            {
                                Good = nowTable.Cell(rowPos, columPos).Range.Text;
                                Good = Good.Remove(Good.Length - 2, 2);

                                if (!(Good.ToString().Equals("") || Good.ToString().Contains("□")))
                                {
                                    switch (columPos)
                                    {
                                        case 1:
                                            material += "~";
                                            material += Good.ToString();
                                            break;
                                        case 2:
                                            material += "!";
                                            material += Good.ToString();
                                            break;
                                        case 3:
                                            material += "@";
                                            material += Good.ToString();
                                            break;
                                        case 4:
                                            material += "#";
                                            material += Good.ToString();
                                            break;
                                        case 5:
                                            material += "$";
                                            material += Good.ToString();
                                            break;
                                        case 6:
                                            material += "%";
                                            material += Good.ToString();
                                            break;
                                        case 7:
                                            material += "^";
                                            material += Good.ToString();
                                            break;
                                        default:
                                            MessageBox.Show("default");
                                            break;


                                    }


                                }
                            }
                            else if (need == 3 & rowsql <= rowPos)
                            {
                                Good = nowTable.Cell(rowPos, columPos).Range.Text;
                                Good = Good.Remove(Good.Length - 2, 2);

                                if (!(Good.ToString().Equals("") || Good.ToString().Contains("□")))
                                {
                                    switch (columPos)
                                    {
                                        case 1:
                                            Device += "~";
                                            Device += Good.ToString();
                                            break;
                                        case 2:
                                            Device += "!";
                                            Device += Good.ToString();
                                            break;
                                        case 3:
                                            Device += "@";
                                            Device += Good.ToString();
                                            break;
                                        case 4:
                                            Device += "#";
                                            Device += Good.ToString();
                                            break;
                                        case 5:
                                            Device += "$";
                                            Device += Good.ToString();
                                            break;
                                        case 6:
                                            Device += "%";
                                            Device += Good.ToString();
                                            break;
                                        case 7:
                                            Device += "^";
                                            Device += Good.ToString();
                                            break;
                                        case 8:
                                            break;
                                        default:
                                            MessageBox.Show("default");
                                            break;


                                    }

                                    // MessageBox.Show(Device);
                                }

                            }


                        }
                        catch
                        { Good = ""; }



                    }


                }
                try
                {
                    int testadd = 0;
                    string taskcardtest = nowTable.Cell(1, 1).Range.Text;

                    taskcardtest = taskcardtest.Remove(taskcardtest.Length - 2, 2);//remove \r\a
                    if (taskcardtest.ToString().Contains("工作說明書") && taskcardtest.ToString().Contains("TASK CARD"))
                    {
                        testadd++;
                    }
                    //   MessageBox.Show(nowTable.Cell(1, 1).Range.Text + "\n" + nowTable.Cell(2, 1).Range.Text + "\n");

                    taskcardtest = taskcardtest.Remove(taskcardtest.Length - 2, 2);//remove \r\a
                    Alltaskcode = nowTable.Cell(1 + testadd, 1).Range.Text;
                    Alltaskcode = Alltaskcode.Remove(Alltaskcode.Length - 2, 2);//remove \r\a
                    Alltaskcode = Alltaskcode.Replace("工作說明書編碼", "").Replace("：", "");
                    Alltaskcode = Alltaskcode.Replace("Task Card Code", "");
                    Alltaskcode = Alltaskcode.TrimStart();

                    Alltaskname = nowTable.Cell(1 + testadd, 2).Range.Text;
                    Alltaskname = Alltaskname.Remove(Alltaskname.Length - 2, 2);//remove \r\a\
                    Alltaskname = Alltaskname.Replace("工作說明書名稱", "").Replace("：", "").Replace("◎SC", "");
                    Alltaskname = Alltaskname.Replace("Task Card Name", "");
                    Alltaskname = Alltaskname.TrimStart();

                    Alltaskdetail = nowTable.Cell(2 + testadd, 1).Range.Text;
                    Alltaskdetail = Alltaskdetail.Remove(Alltaskdetail.Length - 2, 2);//remove \r\a
                    Alltaskdetail = Alltaskdetail.Replace("工作說明書簡述", "").Replace("：", "");
                    Alltaskdetail = Alltaskdetail.Replace("Task Card Description", "");
                    Alltaskdetail = Alltaskdetail.TrimStart();


                    Alltasklocation = nowTable.Cell(3 + testadd, 1).Range.Text;
                    Alltasklocation = Alltasklocation.Remove(Alltasklocation.Length - 2, 2);//remove \r\a
                    Alltasklocation = Alltasklocation.Replace("維修適用範圍", "").Replace("：", "");
                    Alltasklocation = Alltasklocation.Replace("Location", "");
                    Alltasklocation = Alltasklocation.TrimStart();

                    Alltaskdevice = nowTable.Cell(3 + testadd, 2).Range.Text;
                    Alltaskdevice = Alltaskdevice.Remove(Alltaskdevice.Length - 2, 2);//remove \r\a
                    Alltaskdevice = Alltaskdevice.Replace("維修適用設備", "").Replace("：", "");
                    Alltaskdevice = Alltaskdevice.Replace("Equipment Description", "");
                    Alltaskdevice = Alltaskdevice.TrimStart();


                    maintain = nowTable.Cell(3 + testadd, 3).Range.Text;
                    maintain = maintain.Remove(maintain.Length - 2, 2);//remove \r\a
                    maintain = maintain.Replace("維修週期", "").Replace("：", "");
                    maintain = maintain.Replace("Task Interval", "").Replace("Interval", "");
                    maintain = maintain.TrimStart();

                    safe = nowTable.Cell(4 + testadd, 1).Range.Text;
                    safe = safe.Remove(safe.Length - 2, 2);//remove \r\a


                    char[] delimiterChars = { '\n', '\t', '\r', '\a' };//remove \r\a
                    string[] words = safe.Split(delimiterChars);
                    int photoint = 0;
                    safe = "";
                    for (int x = 0; x < words.Length; x++)
                    {

                        if (words[x].ToString().Contains("。"))
                        {

                            photoint++;
                            words[x] = photoint.ToString() + "." + words[x].ToString();
                            safe += words[x] + "\n";


                        }

                    }

                }
                catch (Exception ex)
                { MessageBox.Show(ex.ToString()); }


                BuildFolderTask(Alltaskcode.ToString());
                ///寫進資料庫
                //   string lower = txtCode.Text;//轉換輸入格式
                //   lower = lower.ToLower().ToString();


                MySqlConnection conn = new MySqlConnection(allconnString);//實做一個物件來連線
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                try//測試有沒有表格
                {
                    string showTable = "desc `" + Alltaskcode.ToString().ToLower() + @"`";
                    MySqlCommand Tablefind = new MySqlCommand(showTable, conn);

                    //index1 = Tablefind.ExecuteNonQuery();
                    Tablefind.ExecuteNonQuery();

                    //有表格
                    string update = "UPDATE `taskcard` SET TaskCardName=@test2,TaskLocation=@test3,TaskDetail=@test4,Equipment=@test5,Maintaincycle=@test6,Peopleneed=@test7,Material=@test8,Device= @test9,Safe=@test10 WHERE TaskCardCode = @test1";
                    using (MySqlCommand updatecmd = new MySqlCommand(update, conn))//匯入檔案到資料庫
                    {

                        updatecmd.Parameters.Add("@test1", Alltaskcode.ToString());
                        updatecmd.Parameters.Add("@test2", Alltaskname.ToString());
                        updatecmd.Parameters.Add("@test3", Alltasklocation.ToString());
                        updatecmd.Parameters.Add("@test4", Alltaskdetail.ToString());
                        updatecmd.Parameters.Add("@test5", Alltaskdevice.ToString());
                        updatecmd.Parameters.Add("@test6", maintain.ToString());

                        updatecmd.Parameters.Add("@test7", people.ToString());
                        updatecmd.Parameters.Add("@test8", material.ToString());
                        updatecmd.Parameters.Add("@test9", Device.ToString());
                        updatecmd.Parameters.Add("@test10", safe.ToString());
                        updatecmd.Parameters.Add("@test11", "1");
                        int index = updatecmd.ExecuteNonQuery();
                        bool success = false;
                        if (index > -1)
                            Console.WriteLine(path + "update success");

                    }
                }

                catch //沒有表格
                {

                    string Tablesql = @"CREATE TABLE `" + Alltaskcode.ToString().ToLower() + @"` (TaskCardName VARCHAR(45),Items int(11),TaskCardCode TEXT,TaskDetail TEXT,TaskImage TEXT,TaskImagepath TEXT,StandardValue VARCHAR(45),Note VARCHAR(45),Equipment SET('A','B'))";
                    MySqlCommand creatcmd = new MySqlCommand(Tablesql, conn);
                    int Create = creatcmd.ExecuteNonQuery();


                    ///////注意還未測試
                    string Checksql = @"CREATE TABLE `" + Alltaskcode.ToString().ToLower() + "-f1" + @"` (TaskCardName VARCHAR(45),Items int(11),TaskCardCode TEXT,TaskDetail TEXT,Equipment VARCHAR(45),StandardValue VARCHAR(45),Note VARCHAR(45),CheckResult VARCHAR(45) )";
                    MySqlCommand checkcreatecmd = new MySqlCommand(Checksql, conn);
                    int CheckCreate = checkcreatecmd.ExecuteNonQuery();



                    string Checkrecordsql = @"CREATE TABLE `" + Alltaskcode.ToString().ToLower() + "-f1-record" + @"` (Name VARCHAR(20),Items TEXT,TaskCardCode TEXT,CheckDetail TEXT,CheckResult TEXT,Result ENUM('Qualified','Unqualified',''),Time TIMESTAMP,FinishTime TEXT,TaskLocation ENUM('Zuoying_Station','Tainan_Station','Zuoying_Base','Yanchao_Factory'),Cycle TEXT,Sign VARCHAR(20),Location TEXT,GPS TEXT)";

                    MySqlCommand checkrecordcmd = new MySqlCommand(Checkrecordsql, conn);
                    int Checkrecord = checkrecordcmd.ExecuteNonQuery();



                    string sql = @"INSERT INTO `" + "taskcard" + @"`(`TaskCardCode`,`TaskCardName`,`TaskLocation`,`TaskDetail`,`Equipment`,`Maintaincycle`,`Peopleneed`,`Material`,`Device`,`Safe`,`Version`) VALUES
                           (@test1,@test2,@test3,@test4,@test5,@test6,@test7,@test8,@test9,@test10,@test11) ";
                    using (MySqlCommand cmd = new MySqlCommand(sql, conn))//匯入檔案到資料庫
                    {

                        cmd.Parameters.Add("@test1", Alltaskcode.ToString());
                        cmd.Parameters.Add("@test2", Alltaskname.ToString());
                        cmd.Parameters.Add("@test3", Alltasklocation.ToString());
                        cmd.Parameters.Add("@test4", Alltaskdetail.ToString());
                        cmd.Parameters.Add("@test5", Alltaskdevice.ToString());
                        cmd.Parameters.Add("@test6", maintain.ToString());

                        cmd.Parameters.Add("@test7", people.ToString());
                        cmd.Parameters.Add("@test8", material.ToString());
                        cmd.Parameters.Add("@test9", Device.ToString());
                        cmd.Parameters.Add("@test10", safe.ToString());
                        cmd.Parameters.Add("@test11", "1");
                        int index = cmd.ExecuteNonQuery();
                        bool success = false;
                        if (index > -1)
                            Console.WriteLine(path + "input success");
                    }
                }
                finally
                {
                    doc.Close();

                }

            }


        }
        private void savephoto(string photoname, int count, Microsoft.Office.Interop.Word._Document doc, Microsoft.Office.Interop.Word._Application app)
        {
            //Microsoft.Office.Interop.Word._Application app = new Microsoft.Office.Interop.Word.Application();
            // Microsoft.Office.Interop.Word._Document doc = null;
            //foreach (Microsoft.Office.Interop.Word.InlineShape ish in doc.InlineShapes)
            // {
            // if ((doc.InlineShapes[count].Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapeLinkedPicture) || (doc.InlineShapes[count].Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture))

            //MessageBox.Show(photoname.ToString());
            try
            {
                doc.InlineShapes[count].Select();
            }
            catch { count = count - 2; }
            app.Selection.Copy();
            Image image = Clipboard.GetImage();//取得檔案中圖片
            pictureBox1.Image = image;
            string fs = @"D:\SCD\TaskPhoto\" + photoname.ToString();
            try { pictureBox1.Image.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg); }
            catch { }
            FTPUpload(photoname.ToString());



            // }
        }
        private void Connect(String path)//连接ftp

        {

            // 根据uri创建FtpWebRequest对象

            reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(path));

            // 指定数据传输类型

            reqFTP.UseBinary = true;

            // ftp用户名和密码

            reqFTP.Credentials = new NetworkCredential(ftpUserID, ftpPassword);

        }
        private void BuildFolderTask(string BF)//創建任務資料夾
        {

            FtpWebRequest re = (FtpWebRequest)WebRequest.Create(ftpServerIP + BF.ToString().ToLower());
            re.Credentials = new NetworkCredential(ftpUserID, ftpPassword);
            re.Method = WebRequestMethods.Ftp.MakeDirectory;
            re.Timeout = (60000 * 1);
            try
            {
                FtpWebResponse response = (FtpWebResponse)re.GetResponse();
                response.Close();
            }
            catch { Console.WriteLine("此工卡總表有重複輸入"); }



        }
        private void FTPUpload(string photoname)
        {

            string username = ftpUserID;
            string password = ftpPassword;
            string uploadUrl = ftpServerIP + lower.ToString() + "/" + photoname;

            if (!System.IO.File.Exists(@"D:\SCD\TaskPhoto\" + photoname.ToString()))
                MessageBox.Show("上傳檔案不存在");

            //   MessageBox.Show(i.ToString());
            // 要上傳的檔案
            //圖片格式
            string MyFileName = @"D:\SCD\TaskPhoto\" + photoname.ToString();
            //文件格式擋
            /*StreamReader sourceStream = new StreamReader(@"C:\Users\B510\Desktop\File\" + i.ToString()+".jpg");
            byte[] data = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
            sourceStream.Close();
            */
            WebClient wc = new WebClient();
            wc.Credentials = new NetworkCredential(username, password);
            wc.UploadFile(uploadUrl, MyFileName);
            path = uploadUrl.ToString();
            ///文件格式上傳 wc.UploadData(uploadUrl, data);

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            labTime.Text = DateTime.Now.ToString();

            if (labTime.Text.Contains("下午 01:08:10"))//做資料備份雲端

            {
                allconnString = "server=" + IP.ToString() + ";port=" + port.ToString() + ";user id=thsr2019_05_08;password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫
                string file = @"D:\SCD\SQLBackup\SQLbackup_" + DateTime.Now.ToString("MM_dd_yyyy") + ".sql";//資料庫備份本機端路徑
                using (MySqlConnection conn = new MySqlConnection(allconnString))//實做一個物件
                {
                    if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                    using (MySqlCommand cmd = new MySqlCommand())
                    {
                        using (MySqlBackup mb = new MySqlBackup(cmd))
                        {
                            try
                            {
                                cmd.Connection = conn;
                                // conn.Open();
                                mb.ExportToFile(file);
                                conn.Close();
                                Console.WriteLine("SQLbackup    =  Done");
                                //傳爆FTP上
                                string username = ftpUserID;
                                string password = ftpPassword;
                                string uploadUrl = ftpServerIP + "資料庫備份/" + DateTime.Now.ToString("MM_dd_yyyy") + ".sql";


                                if (!System.IO.File.Exists(@"D:\SCD\SQLBackup\SQLbackup_" + DateTime.Now.ToString("MM_dd_yyyy") + ".sql"))
                                    MessageBox.Show("上傳檔案不存在");
                                // 要上傳的檔案               
                                else
                                {
                                    string MyFileName = @"D:\SCD\SQLBackup\SQLbackup_" + DateTime.Now.ToString("MM_dd_yyyy") + ".sql";

                                    WebClient wc = new WebClient();
                                    wc.Credentials = new NetworkCredential(username, password);
                                    wc.UploadFile(uploadUrl, MyFileName);
                                    Console.WriteLine("成功 ， 儲存的路徑為=" + uploadUrl.ToString());
                                }


                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.ToString());
                            }
                        }

                    }
                }
            }
            else// if (labTime.Text.Contains("下午 03::50"))//word匯入
            {
                int xt = 1;

                foreach (Process item in Process.GetProcessesByName("WINWORD"))
                { item.Kill(); }//如存在開啟的Word則先關閉(一個應用程式例項就是程序)
                MySqlConnection conn = new MySqlConnection(allconnString);//實做一個物件來連線
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                foreach (string fname in System.IO.Directory.GetFileSystemEntries(@"D:\SCD\下載\TaskDownload\"))
                {

                    Console.WriteLine(fname);
                    count = 1;
                    DeleteFileName(fname.Remove(0, 22));//FTP 上刪除
                    path = fname;//傳給TASKCARD
                    Taskcard();

                    Microsoft.Office.Interop.Word._Application app = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word._Document doc = null;

                    object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument;
                    object unknow = Type.Missing;

                    object oFileName = @"" + fname + "";


                    doc = app.Documents.Open(ref oFileName, ref format,//資料路徑 格式
                           ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow);
                    //input工卡總表





                    ///抓WORD表格

                    for (int tablePos = 2; tablePos <= 3; tablePos++)//抓第幾個表格
                    {

                        if (tablePos == 2)
                        {
                            string name = "";

                            int i = 1;
                            Microsoft.Office.Interop.Word.Table nowTable = doc.Tables[2];
                            int testadd = 0;
                            string taskcardtest = nowTable.Cell(1, 1).Range.Text;

                            taskcardtest = taskcardtest.Remove(taskcardtest.Length - 2, 2);//remove \r\a
                                                                                           // MessageBox.Show(taskcardtest.ToString());
                            if ((taskcardtest.ToString().Contains("工作說明書") && taskcardtest.ToString().Contains("TASK CARD")) || (taskcardtest.ToString().Contains("工作說明書") && taskcardtest.ToString().Contains("Task Card") && !(taskcardtest.ToString().Contains("編碼"))))

                            {
                                //  MessageBox.Show("rowPos+1");
                                testadd = 1;
                            }


                            for (int rowPos = 1 + testadd; rowPos <= nowTable.Rows.Count; rowPos++)//抓第幾----
                            {
                                Detail1 = "";

                                for (int columPos = 1; columPos <= nowTable.Columns.Count; columPos++)//抓底幾個|||
                                {



                                    try
                                    {
                                        Good = nowTable.Cell(rowPos, columPos).Range.Text;
                                        Good = Good.Remove(Good.Length - 2, 2);//remove \r\a
                                                                               //MessageBox.Show("Rowpos = "+rowPos + "\ncolumPos = " + columPos +"\n"+ Good.ToString());
                                    }
                                    catch
                                    { //MessageBox.Show(columPos.ToString()+"        "+rowPos);
                                    }



                                    if (Good.ToString().Equals("N/A"))
                                        Good = "";
                                    switch (columPos)
                                    {
                                        case 1:
                                            Item = Good.TrimStart().TrimEnd();

                                            if (rowPos == 1 + testadd)
                                            {//取TaskCardCode
                                             // MessageBox.Show("Rowpos = " + rowPos + "\ncolumPos = " + columPos + "\n" + Item.ToString());
                                                Item = Item.Replace("工作說明書編碼", "").Replace("：", "").Replace("Task Card Name", "");
                                                Item = Item.Replace("Task Card Code", "");
                                                Item = Item.TrimStart().TrimEnd();

                                                lower = Item.ToString();
                                                // MessageBox.Show("Rowpos = " + rowPos + "\ncolumPos = " + columPos + "\n" + lower.ToString());
                                                //刪除TASKCAR資料
                                                string delete = @"DELETE FROM`" + lower.ToString() + "`";
                                                MySqlCommand delcmd = new MySqlCommand(delete, conn);
                                                index = delcmd.ExecuteNonQuery();
                                                delcmd.Dispose();
                                                //刪除checklist資料   
                                                string deletecheck = @"DELETE FROM`" + lower.ToString() + "-f1`";
                                                MySqlCommand delcmdcheck = new MySqlCommand(deletecheck, conn);
                                                index = delcmdcheck.ExecuteNonQuery();
                                                delcmdcheck.Dispose();


                                            }
                                            break;
                                        case 2:
                                            Detail = Good;
                                            if (rowPos == 1 + testadd)
                                            {//取工作內容
                                                Detail = Detail.Replace("工作說明書名稱", "").Replace("：", "").Replace("◎SC", "");
                                                Detail = Detail.Replace("Task Card Name", "");
                                                Detail = Detail.TrimStart();

                                                nametask = Detail.ToString();//轉換輸入格式

                                            }
                                            else
                                            {//資料整理
                                                char[] delimiterChars = { '\n', '\t', '\r', '\a' };//remove \r\a
                                                name = Detail.ToString();//轉換輸入格式
                                                string[] words = name.Split(delimiterChars);
                                                int photoint = 0;
                                                for (int x = 0; x < words.Length; x++)
                                                {
                                                    // Console.WriteLine(words[x].ToString() + "長度為" + words.Length + "現在執行到" + x);
                                                    //標題項目重新
                                                    //以下是偵測規0
                                                    if (words[x].ToString().Contains("取得") && words[x].ToString().Contains("權限"))
                                                        photoint = 0;


                                                    //else if (words[x].ToString().Substring(words[x].Length - 1, 1).Equals("。")|| (words[x].ToString().Contains("。") && words[x].ToString().Substring(words[x].Length - 1, 1).Equals(")"))|| words[x].ToString().Substring(words[x].Length - 1, 1).Equals("↵"))
                                                    if (words[x].ToString().Contains("。"))
                                                    {
                                                        photoint++;
                                                        words[x] = photoint.ToString() + "." + words[x].ToString();
                                                        // MessageBox.Show(words[x].ToString(),"Test",MessageBoxButtons.OKCancel ,MessageBoxIcon.Information);

                                                    }
                                                    if (words[x].ToString().Equals("/ //"))
                                                    {
                                                        for (int twophoto = 1; twophoto <= 3; twophoto++)
                                                        {
                                                            photoname = lower.ToString() + "_" + i.ToString() + ".jpg";

                                                            if (twophoto >= 2)
                                                                words[x] += photoname.ToString();
                                                            else
                                                                words[x] = photoname.ToString();
                                                            //MessageBox.Show(words[x].ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                            savephoto(photoname, count, doc, app);
                                                            i++;
                                                            count++;
                                                        }
                                                    }
                                                    else if (words[x].ToString().Contains("上 /、下 / 鍵或左 /、右 / 鍵"))
                                                    {
                                                        for (int twophoto = 1; twophoto <= 4; twophoto++)
                                                        {
                                                            photoname = lower.ToString() + "_" + i.ToString() + ".jpg";
                                                            switch (twophoto)
                                                            {
                                                                case 1:
                                                                    words[x] = words[x].ToString().Replace("上 /", "上" + photoname.ToString());
                                                                    break;
                                                                case 2:
                                                                    words[x] = words[x].ToString().Replace("下 / ", "下" + photoname.ToString());
                                                                    break;
                                                                case 3:
                                                                    words[x] = words[x].ToString().Replace("左 /", "左" + photoname.ToString());
                                                                    break;
                                                                case 4:
                                                                    words[x] = words[x].ToString().Replace("右 /", "右" + photoname.ToString());
                                                                    break;

                                                            }



                                                            savephoto(photoname, count, doc, app);
                                                            i++;
                                                            count++;
                                                        }
                                                    }
                                                    else if (words[x].ToString().Contains("「雨刷」按鈕") && words[x].ToString().Contains("確認雨刷動作正常後須點選按鈕"))
                                                    {
                                                        for (int twophoto = 1; twophoto <= 2; twophoto++)
                                                        {
                                                            photoname = lower.ToString() + "_" + i.ToString() + ".jpg";
                                                            switch (twophoto)
                                                            {
                                                                case 1:
                                                                    words[x] = words[x].ToString().Replace("「雨刷」", "「雨刷」" + photoname.ToString());
                                                                    break;
                                                                case 2:
                                                                    words[x] = words[x].ToString().Replace("確認雨刷動作正常後須點選", "確認雨刷動作正常後須點選" + photoname.ToString());
                                                                    break;

                                                            }

                                                            savephoto(photoname, count, doc, app);
                                                            i++;
                                                            count++;
                                                        }
                                                    }

                                                    else if (words[x].ToString().Equals("//") || words[x].ToString().Contains("/ /") || words[x].ToString().Contains("/  /") || words[x].ToString().Contains("/   /") || words[x].ToString().Contains("/     /") || words[x].ToString().Equals("/             /")
                                                        || words[x].ToString().Equals("/ ") || words[x].ToString().Equals(" /") || words[x].ToString().Equals("/    /"))
                                                    {
                                                        for (int twophoto = 1; twophoto <= 2; twophoto++)
                                                        {
                                                            photoname = lower.ToString() + "_" + i.ToString() + ".jpg";

                                                            if (twophoto == 2)
                                                                words[x] += photoname.ToString();
                                                            else
                                                                words[x] = photoname.ToString();
                                                            // MessageBox.Show(words[x].ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                            savephoto(photoname, count, doc, app);
                                                            i++;
                                                            count++;
                                                        }


                                                    }
                                                    /// \ \                                   \  \                                  \   \                                   \    \              
                                                    else if (words[x].ToString().Contains(" ") || words[x].ToString().Contains("  ") || words[x].ToString().Contains("   ") || words[x].ToString().Contains("    ") || words[x].ToString().Contains("    ") || words[x].ToString().Contains("    "))
                                                    {
                                                        for (int twophoto = 1; twophoto <= 2; twophoto++)
                                                        {
                                                            photoname = lower.ToString() + "_" + i.ToString() + ".jpg";

                                                            if (twophoto == 2)
                                                                words[x] += photoname.ToString();
                                                            else
                                                                words[x] = photoname.ToString();
                                                          

                                                            savephoto(photoname, count, doc, app);
                                                            i++;
                                                            count++;
                                                        }


                                                    }
                                                    else if (words[x].ToString().Contains("圖示/") || words[x].ToString().Contains("OK /") || words[x].ToString().Equals("/") || words[x].ToString().Equals(" /") || words[x].ToString().Equals("  /") || words[x].ToString().Equals("   /") || words[x].ToString().Equals("    /")
                                                                   || words[x].ToString().Equals("/ "))
                                                    // else if(words[x].ToString().Equals("/"))
                                                    {

                                                        if (words[x].ToString().Contains("圖示/") || words[x].ToString().Contains("OK /"))
                                                        {
                                                            photoname = lower.ToString() + "_" + i.ToString() + ".jpg";
                                                            words[x] = words[x].ToString().Replace("/", photoname.ToString());
                                                        }


                                                        else
                                                        {
                                                            photoname = lower.ToString() + "_" + i.ToString() + ".jpg";
                                                            words[x] = photoname.ToString();
                                                        }


                                                        //  Console.WriteLine(photoname.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                        savephoto(words[x], count, doc, app);
                                                        i++;
                                                        count++;

                                                    }


                                                    else if (words[x].ToString().Equals("") || words[x].ToString().Contains("   ") || words[x].ToString().Equals(" ") || words[x].ToString().Equals("  ") || words[x].ToString().Equals("   ") || words[x].ToString().Equals("    "))
                                                    // else if(words[x].ToString().Equals("/"))
                                                    {


                                                        photoname = lower.ToString() + "_" + i.ToString() + ".jpg";

                                                        words[x] = photoname.ToString();
                                                        //  MessageBox.Show(words[x].ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                        savephoto(words[x], count, doc, app);
                                                        i++;
                                                        count++;

                                                    }


                                                    if (rowPos >= 3)
                                                    {
                                                        if (words[x].ToString().Equals(""))
                                                            Detail1 += words[x].ToString();
                                                        else
                                                            Detail1 += words[x].ToString() + "\n";
                                                        // MessageBox.Show(Detail1.ToString());
                                                    }
                                                }



                                            }
                                            break;
                                        case 4:
                                            standard = Good;
                                            break;
                                        case 5:
                                            Note = Good;
                                            break;
                                    }





                                }//欄位


                                ////寫進資料庫


                                if (rowPos >= 3 + testadd && !(Item.ToString().Contains("作業編號")))
                                {


                                    if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

                                    string sql = @"INSERT INTO `" + lower.ToString() + @"`(`TaskCardName`,`Items`,`TaskCardCode`,`TaskDetail`,`StandardValue`,`Note`,`TaskImage`,`TaskImagepath`) VALUES
                               (@test1,@test2,@test3,@test4,@test5,@test6,@Image,@Imagepath) ";

                                    using (MySqlCommand cmd = new MySqlCommand(sql, conn))//匯入檔案到資料庫
                                    {

                                        cmd.Parameters.Add("@test1", nametask.ToString());
                                        cmd.Parameters.Add("@test2", Item.ToString());
                                        cmd.Parameters.Add("@test3", lower.ToString().ToLower());

                                        if (rowPos >= 3)
                                            cmd.Parameters.Add("@test4", Detail1.ToString());
                                        else
                                            cmd.Parameters.Add("@test4", Detail.ToString());
                                        cmd.Parameters.Add("@test5", standard.ToString());
                                        cmd.Parameters.Add("@test6", Note.ToString());
                                        cmd.Parameters.Add("@Image", lower.ToString() + "_" + Item.ToString());
                                        cmd.Parameters.Add("@Imagepath", ftpServerIP + lower.ToString() + "_" + Item.ToString());
                                        index = cmd.ExecuteNonQuery();
                                        cmd.Dispose();
                                        if (index > 0)
                                            Console.WriteLine("TASK匯入OK");
                                    }





                                }//資料庫

                            }//每列



                        }//TaskCard匯入if
                         //CheckList匯入
                        if (tablePos == 3)
                        {

                            int i = 1;
                            //崩潰時站存資料
                            string colum = "";
                            string CC = "";
                            string tt = "";
                            string ch = "";
                            string st = "";
                            string no = "";

                            int changindex = 1;//轉換格式
                            int index = 0;//判斷資料有無會盡
                                          //傳到資料庫
                            string Item = "";
                            string equipment = "";
                            string taskdetail = "";
                            string checkresult = "";
                            string stand = "";
                            string note = "";
                            string maybenote = "";
                            string maybeone = "";

                            //抓資料格式
                            int a1 = 1;
                            int a2 = 2;
                            int a3 = 3;
                            int a4 = 4;
                            int a5 = 5;
                            int a6 = 6;
                            int a7 = 7;
                            int a8 = 8;
                            int a9 = 9;
                            try
                            {
                                Microsoft.Office.Interop.Word.Table nowTable = doc.Tables[4];

                                for (int rowPos = 3; rowPos <= nowTable.Rows.Count; rowPos++)//抓第幾----
                                {
                                    int title = 0;
                                    int outland = 0;
                                    int change9_6 = 0;
                                    string[] allinput = new string[9];
                                    for (int columPos = 1; columPos <= nowTable.Columns.Count; columPos++)//抓底幾個|||

                                    {

                                        switch (nowTable.Columns.Count)
                                        {
                                            case 6:
                                                Console.WriteLine("這是6格");

                                                if (rowPos >= 3)
                                                {
                                                    try
                                                    {

                                                        Good = nowTable.Cell(rowPos, columPos).Range.Text;//多出來崩饋點
                                                        Good = Good.Remove(Good.Length - 2, 2);
                                                        Console.WriteLine(Good.ToString());

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                colum = Good.ToString();//崩潰
                                                                Item = Good.ToString().TrimStart().TrimEnd();

                                                                break;
                                                            case 2:
                                                                CC = Good.ToString();//崩潰
                                                                equipment = Good.ToString();

                                                                if (equipment.ToString().Equals("完成檢查作業"))
                                                                {
                                                                    tt = ""; ch = ""; st = ""; no = "";
                                                                }
                                                                break;
                                                            case 3:
                                                                if (Good.ToString().Equals("檢查項目"))
                                                                {
                                                                    title++;
                                                                }
                                                                tt = Good.ToString();//崩潰
                                                                taskdetail = Good.ToString();

                                                                break;

                                                            case 6:
                                                                no = Good.ToString();//崩潰
                                                                note = Good.ToString();
                                                                // MessageBox.Show(note.ToString());
                                                                break;

                                                        }
                                                    }
                                                    catch
                                                    {

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                Item = colum.ToString();
                                                                break;
                                                            case 2:
                                                                equipment = CC.ToString();
                                                                //MessageBox.Show(equipment.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 3:
                                                                taskdetail = tt.ToString();
                                                                // MessageBox.Show(taskdetail.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 6:
                                                                note = no.ToString();
                                                                //  MessageBox.Show(note.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;

                                                        }



                                                    }




                                                }//if判斷
                                                break;
                                            case 7:
                                                Console.WriteLine("這是7格");
                                                if (rowPos >= 3)
                                                {
                                                    try
                                                    {

                                                        Good = nowTable.Cell(rowPos, columPos).Range.Text;//多出來崩饋點
                                                        Good = Good.Remove(Good.Length - 2, 2);

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                if (Good.ToString().Equals("下列檢查項目適用於 「室外型」 CAF設備"))
                                                                    title++;
                                                                colum = Good.ToString();//崩潰
                                                                Item = Good.ToString();

                                                                break;
                                                            case 2:
                                                                CC = Good.ToString();//崩潰
                                                                equipment = Good.ToString();

                                                                if (equipment.ToString().Equals("完成檢查作業"))
                                                                {
                                                                    tt = ""; ch = ""; st = ""; no = "";
                                                                }
                                                                break;
                                                            case 3:
                                                                if (Good.ToString().Equals("檢查項目") || Good.ToString().Equals("檢 查 標 準"))
                                                                {
                                                                    title++;
                                                                }
                                                                tt = Good.ToString();//崩潰
                                                                taskdetail = Good.ToString();

                                                                break;
                                                            case 4:
                                                                if (Good.ToString().Equals("檢查標準") || Good.ToString().Equals("檢 查 標 準"))
                                                                {
                                                                    title++;
                                                                }
                                                                st = Good.ToString();//崩潰
                                                                stand = Good.ToString();
                                                                if (st.ToString().Equals("□"))
                                                                {
                                                                    st = ""; stand = "";
                                                                }


                                                                break;
                                                            case 6://7對6表格

                                                                if (Good.ToString().Equals("□"))
                                                                {

                                                                }
                                                                else
                                                                {
                                                                    no = Good.ToString();//崩潰
                                                                    note = Good.ToString();
                                                                    // MessageBox.Show(note.ToString());
                                                                }
                                                                break;

                                                            case 7:
                                                                no = Good.ToString();//崩潰
                                                                note = Good.ToString();
                                                                // MessageBox.Show(note.ToString());
                                                                break;

                                                        }
                                                    }
                                                    catch
                                                    {

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                Item = colum.ToString();
                                                                break;
                                                            case 2:
                                                                equipment = CC.ToString();
                                                                //MessageBox.Show(equipment.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 3:
                                                                taskdetail = tt.ToString();
                                                                // MessageBox.Show(taskdetail.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 4:
                                                                stand = st.ToString();
                                                                break;
                                                            case 6:
                                                                note = no.ToString();
                                                                break;
                                                            case 7:
                                                                note = no.ToString();
                                                                //  MessageBox.Show(note.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;

                                                        }



                                                    }




                                                }//if判斷
                                                break;

                                            case 8:
                                                Console.WriteLine("這是8格");
                                                if (rowPos >= 3)
                                                {
                                                    try
                                                    {

                                                        Good = nowTable.Cell(rowPos, columPos).Range.Text;//多出來崩饋點
                                                        Good = Good.Remove(Good.Length - 2, 2);
                                                        //    MessageBox.Show(Good.ToString());
                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                colum = Good.ToString();//崩潰
                                                                Item = Good.ToString();

                                                                break;
                                                            case 2:
                                                                CC = Good.ToString();//崩潰
                                                                equipment = Good.ToString();

                                                                if (equipment.ToString().Equals("完成檢查作業"))
                                                                {
                                                                    tt = ""; ch = ""; st = ""; no = "";
                                                                }
                                                                break;
                                                            case 3:
                                                                if (Good.ToString().Equals("檢查項目") || Good.ToString().Equals("檢 查 項 目"))
                                                                {
                                                                    title++;//讓這行不進行匯入
                                                                }
                                                                tt = Good.ToString();//崩潰
                                                                taskdetail = Good.ToString();

                                                                break;
                                                            case 4:
                                                                ch = Good.ToString();//崩潰
                                                                checkresult = Good.ToString().Replace("　", "");
                                                                if (ch.ToString().Equals("□"))
                                                                {
                                                                    ch = ""; checkresult = "";
                                                                }

                                                                break;
                                                            case 5:
                                                                if (Good.ToString().Equals("合格標準"))
                                                                {
                                                                    title++;
                                                                }
                                                                st = Good.ToString();//崩潰
                                                                stand = Good.ToString();
                                                                if (st.ToString().Equals("□"))
                                                                {
                                                                    st = ""; stand = "";
                                                                }
                                                                break;
                                                            case 6://8表格對6NOTE

                                                                if (Good.ToString().Equals("□"))
                                                                {

                                                                }
                                                                else
                                                                {
                                                                    no = Good.ToString();
                                                                    note = Good.ToString();
                                                                }
                                                                break;

                                                            case 8:
                                                                no = Good.ToString();//崩潰
                                                                note = Good.ToString();

                                                                // MessageBox.Show(note.ToString());
                                                                break;

                                                        }
                                                    }
                                                    catch
                                                    {

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                Item = colum.ToString();
                                                                break;
                                                            case 2:
                                                                equipment = CC.ToString();
                                                                //MessageBox.Show(equipment.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 3:
                                                                taskdetail = tt.ToString();
                                                                // MessageBox.Show(taskdetail.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 4:
                                                                checkresult = ch.ToString().Replace("　", "");
                                                                //  MessageBox.Show(checkresult.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 5:
                                                                stand = st.ToString();
                                                                // MessageBox.Show(stand, "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 6:
                                                                note = no.ToString();
                                                                break;
                                                            case 8:
                                                                note = no.ToString();
                                                                //MessageBox.Show(note.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;

                                                        }



                                                    }




                                                }//if判斷
                                                break;
                                            case 9: //蘭位數9格
                                                Console.WriteLine("這是9格");
                                                if (rowPos >= 3)
                                                {

                                                    try
                                                    {

                                                        Good = nowTable.Cell(rowPos, columPos).Range.Text;//多出來崩饋點
                                                        Good = Good.Remove(Good.Length - 2, 2);

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                colum = Good.ToString();//崩潰
                                                                Item = Good.ToString();

                                                                break;
                                                            case 2:
                                                                CC = Good.ToString();//崩潰
                                                                equipment = Good.ToString();

                                                                if (equipment.ToString().Equals("完成檢查作業"))
                                                                {
                                                                    tt = ""; ch = ""; st = ""; no = "";
                                                                }
                                                                break;
                                                            case 3:
                                                                if (Good.ToString().Equals("檢查項目"))
                                                                {
                                                                    title++;
                                                                }
                                                                tt = Good.ToString();//崩潰
                                                                taskdetail = Good.ToString();
                                                                if (taskdetail.ToString().Equals("記錄項目"))
                                                                {
                                                                    // title++;
                                                                    changindex = 2;
                                                                }
                                                                break;
                                                            case 4:

                                                                ch = Good.ToString();//崩潰
                                                                checkresult = Good.ToString().Replace("　", "");

                                                                break;
                                                            case 5:
                                                                if (Good.ToString().Equals("檢查結果"))
                                                                {
                                                                    title++;
                                                                }

                                                                st = Good.ToString();//崩潰
                                                                stand = Good.ToString();
                                                                break;
                                                            case 6:
                                                                if (Good.ToString().Equals("合格標準"))
                                                                {
                                                                    title++;
                                                                }
                                                                maybeone = Good.ToString();//崩潰
                                                                maybeone = Good.ToString();
                                                                break;

                                                            case 8:
                                                                no = Good.ToString();//崩潰
                                                                note = Good.ToString();
                                                                break;
                                                            case 9:
                                                                if (changindex == 2)//當為記錄時
                                                                {
                                                                    maybenote = Good.ToString();
                                                                    note = maybenote;
                                                                    checkresult = checkresult + "\t" + stand + "\t" + maybeone;
                                                                    stand = "";
                                                                }
                                                                else if (changindex == 1)//表格蘭有9時
                                                                {
                                                                    maybenote = Good.ToString();
                                                                    note = maybenote;
                                                                    taskdetail = taskdetail + checkresult;
                                                                    checkresult = stand;
                                                                    stand = maybeone;

                                                                }

                                                                break;

                                                        }
                                                    }
                                                    catch
                                                    {

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                Item = colum.ToString();
                                                                break;
                                                            case 2:
                                                                equipment = CC.ToString();

                                                                break;
                                                            case 3:
                                                                taskdetail = tt.ToString();
                                                                // MessageBox.Show(taskdetail.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 4:
                                                                checkresult = ch.ToString().Replace("　", "");
                                                                //  MessageBox.Show(checkresult.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 5:
                                                                stand = st.ToString();
                                                                // MessageBox.Show(stand, "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 7:     //9對6
                                                                if (stand.ToString().Equals("□"))
                                                                {
                                                                    stand = "";
                                                                    //MessageBox.Show(maybeone.ToString());
                                                                    note = maybeone.ToString();
                                                                    change9_6 = 1;
                                                                }


                                                                break;
                                                            case 8:
                                                                if (change9_6 == 0)
                                                                {
                                                                    note = no.ToString();
                                                                }


                                                                //  MessageBox.Show(note.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 9:
                                                                if (!(maybeone.ToString().Equals("□") || stand.ToString().Contains("僅紀錄")))
                                                                {
                                                                    if (change9_6 == 0)
                                                                    {
                                                                        note = maybenote;
                                                                        taskdetail = taskdetail + checkresult;
                                                                        checkresult = stand;
                                                                        stand = maybeone;
                                                                    }

                                                                }

                                                                break;


                                                        }


                                                    }




                                                }//if判斷
                                                break;
                                            case 10:
                                                Console.WriteLine("這是10格");
                                                if (rowPos >= 3)
                                                {
                                                    try
                                                    {

                                                        Good = nowTable.Cell(rowPos, columPos).Range.Text;//多出來崩饋點
                                                        Good = Good.Remove(Good.Length - 2, 2);

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                if (Good.ToString().Equals("下列檢查項目適用於 「室外型」 CAF設備"))
                                                                    title++;
                                                                colum = Good.ToString();//崩潰
                                                                Item = Good.ToString();

                                                                break;
                                                            case 2:
                                                                CC = Good.ToString();//崩潰
                                                                equipment = Good.ToString();

                                                                if (equipment.ToString().Equals("完成檢查作業"))
                                                                {
                                                                    tt = ""; ch = ""; st = ""; no = "";
                                                                }
                                                                break;
                                                            case 3:
                                                                if (Good.ToString().Equals("檢查項目") || Good.ToString().Equals("檢 查 標 準"))
                                                                {
                                                                    title++;
                                                                }
                                                                tt = Good.ToString();//崩潰
                                                                taskdetail = Good.ToString();

                                                                break;
                                                            case 4:
                                                                if (Good.ToString().Equals("檢查標準") || Good.ToString().Equals("檢 查 標 準"))
                                                                {
                                                                    title++;
                                                                }
                                                                st = Good.ToString();//崩潰
                                                                stand = Good.ToString();
                                                                if (st.ToString().Equals("□"))
                                                                {
                                                                    st = ""; stand = "";
                                                                }


                                                                break;
                                                            case 6://7對6表格

                                                                if (Good.ToString().Equals("□"))
                                                                {

                                                                }
                                                                else
                                                                {
                                                                    no = Good.ToString();//崩潰
                                                                    note = Good.ToString();
                                                                    // MessageBox.Show(note.ToString());
                                                                }
                                                                break;

                                                            case 7:
                                                                no = Good.ToString();//崩潰
                                                                note = Good.ToString();
                                                                // MessageBox.Show(note.ToString());
                                                                break;

                                                        }
                                                    }
                                                    catch
                                                    {

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                Item = colum.ToString();
                                                                break;
                                                            case 2:
                                                                equipment = CC.ToString();
                                                                //MessageBox.Show(equipment.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 3:
                                                                taskdetail = tt.ToString();
                                                                // MessageBox.Show(taskdetail.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 4:
                                                                stand = st.ToString();
                                                                break;
                                                            case 6:
                                                                note = no.ToString();
                                                                break;
                                                            case 7:
                                                                note = no.ToString();
                                                                //  MessageBox.Show(note.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 10:
                                                                if (stand.ToString().Equals("正 常"))
                                                                { stand = ""; note = ""; }
                                                                else if (note.ToString().Contains("異 常") || note.ToString().Equals("□"))
                                                                    note = "";
                                                                break;

                                                        }



                                                    }




                                                }//if判斷
                                                break;
                                            case 11:
                                                Console.WriteLine("這是11格");
                                                if (rowPos >= 3)
                                                {
                                                    try
                                                    {

                                                        Good = nowTable.Cell(rowPos, columPos).Range.Text;//多出來崩饋點
                                                        Good = Good.Remove(Good.Length - 2, 2);
                                                        //    MessageBox.Show(Good.ToString());
                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                colum = Good.ToString();//崩潰
                                                                Item = Good.ToString();

                                                                break;
                                                            case 2:
                                                                CC = Good.ToString();//崩潰
                                                                equipment = Good.ToString();

                                                                if (equipment.ToString().Equals("完成檢查作業"))
                                                                {
                                                                    tt = ""; ch = ""; st = ""; no = "";
                                                                }
                                                                break;
                                                            case 3:
                                                                if (Good.ToString().Equals("檢查項目") || Good.ToString().Equals("檢 查 項 目"))
                                                                {
                                                                    title++;//讓這行不進行匯入
                                                                }
                                                                tt = Good.ToString();//崩潰
                                                                taskdetail = Good.ToString();

                                                                break;
                                                            case 4:
                                                                ch = Good.ToString();//崩潰
                                                                checkresult = Good.ToString().Replace("　", "");
                                                                if (ch.ToString().Equals("□"))
                                                                {
                                                                    ch = ""; checkresult = "";
                                                                }

                                                                break;
                                                            case 5:
                                                                if (Good.ToString().Equals("合格標準"))
                                                                {
                                                                    title++;
                                                                }
                                                                st = Good.ToString();//崩潰
                                                                stand = Good.ToString();
                                                                if (st.ToString().Equals("□"))
                                                                {
                                                                    st = ""; stand = "";
                                                                }
                                                                break;
                                                            case 6://8表格對6NOTE

                                                                if (Good.ToString().Equals("□"))
                                                                {

                                                                }
                                                                else
                                                                {
                                                                    no = Good.ToString();
                                                                    note = Good.ToString();
                                                                }
                                                                break;

                                                            case 8:
                                                                no = Good.ToString();//崩潰
                                                                note = Good.ToString();

                                                                // MessageBox.Show(note.ToString());
                                                                break;

                                                        }
                                                    }
                                                    catch
                                                    {

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                Item = colum.ToString();
                                                                break;
                                                            case 2:
                                                                equipment = CC.ToString();
                                                                //MessageBox.Show(equipment.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 3:
                                                                taskdetail = tt.ToString();
                                                                // MessageBox.Show(taskdetail.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 4:
                                                                checkresult = ch.ToString().Replace("　", "");
                                                                //  MessageBox.Show(checkresult.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 5:
                                                                stand = st.ToString();
                                                                // MessageBox.Show(stand, "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 6:
                                                                note = no.ToString();
                                                                break;
                                                            case 8:
                                                                note = no.ToString();
                                                                //MessageBox.Show(note.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 11:
                                                                if (stand.ToString().Equals("V"))
                                                                    stand = ""; note = "";
                                                                if (Note.ToString().Equals("□"))
                                                                    note = "";
                                                                break;

                                                        }



                                                    }




                                                }//if判斷
                                                break;
                                            default:
                                                Console.WriteLine("這是每日");
                                                note = "每日檢查";
                                                if (rowPos >= 3)
                                                {
                                                    try
                                                    {

                                                        Good = nowTable.Cell(rowPos, columPos).Range.Text;//多出來崩饋點
                                                        Good = Good.Remove(Good.Length - 2, 2);
                                                        //    MessageBox.Show(Good.ToString());
                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                colum = Good.ToString();//崩潰
                                                                Item = Good.ToString(); break;
                                                            case 2:
                                                                CC = Good.ToString();//崩潰
                                                                equipment = Good.ToString();
                                                                break;
                                                            case 3:
                                                                tt = Good.ToString();//崩潰
                                                                taskdetail = Good.ToString();

                                                                break;
                                                        }
                                                    }
                                                    catch
                                                    {

                                                        switch (columPos)
                                                        {
                                                            case 1:
                                                                Item = colum.ToString();
                                                                break;
                                                            case 2:
                                                                equipment = CC.ToString();
                                                                //MessageBox.Show(equipment.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;
                                                            case 3:
                                                                taskdetail = tt.ToString();
                                                                // MessageBox.Show(taskdetail.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                                break;


                                                        }



                                                    }




                                                }//if判斷
                                                break;
                                        }


                                    }//每||| 
                                     ///寫進資料庫
                                    if (title == 0)
                                    {
                                        if (!(equipment.ToString().Equals("")))
                                        {
                                            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                                            string sql = @"INSERT INTO `" + lower.ToString().ToLower() + "-f1" + @"`(`TaskCardCode`,`TaskCardName`,`Items`,`TaskDetail`,`Equipment`,`CheckResult`,`StandardValue`,`Note`) VALUES
                                                       (@test1,@test2,@test3,@test4,@test5,@test6,@test7,@test8) ";
                                            using (MySqlCommand cmd = new MySqlCommand(sql, conn))//匯入檔案到資料庫
                                            {

                                                cmd.Parameters.Add("@test1", lower.ToString().ToLower());
                                                cmd.Parameters.Add("@test2", nametask.ToString());
                                                cmd.Parameters.Add("@test3", Item.ToString());//Item
                                                cmd.Parameters.Add("@test5", equipment.ToString());
                                                cmd.Parameters.Add("@test4", taskdetail.ToString());

                                                cmd.Parameters.Add("@test6", checkresult.ToString());

                                                cmd.Parameters.Add("@test7", stand.ToString());
                                                cmd.Parameters.Add("@test8", note.ToString());

                                                index = cmd.ExecuteNonQuery();
                                                if (index > 0)
                                                    Console.WriteLine("CHECK匯入OK");

                                            }
                                            Item = ""; equipment = ""; taskdetail = ""; checkresult = ""; stand = ""; note = "";

                                        }

                                    }
                                    title = 0;



                                }//每列
                            }
                            catch (Exception ex) { MessageBox.Show(ex.ToString() + "此表單沒有CHECKLIST"); }
                        }


                    }//抓表格 第幾個表格
                    if (index > 0)
                        Console.WriteLine("匯入成功", "資料庫提醒", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);

                    doc.Close();

                }//每筆資料都匯入
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
                    Console.WriteLine("刪除成功");
                }
                //刪除FTP 上的檔案



            }
        }
    }
}
