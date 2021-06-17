using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZXing;
using ZXing.QrCode;

namespace Test
{
    public partial class Form2 :MetroFramework.Forms.MetroForm
    {//工卡總表
        public string IP = "";
        public string port = "";
        string Good = "";//記錄資料的
        string path = "";
        string filename = "";
        //排程匯入資料用的
        string excelpath = "";
        string excelname = "";
        // 訪只重複輸入到帳號權限GRID
        int i = 1;
        int i2 = 1;//防止ＲＥ新增資料
        public Form1 f1 = null;
        //抓帳戶資料所需要的
        string permission = "";
       public  int style = 0;

        string connString = "";
        public Form2(string name,string team,string permission,int style)
        {   InitializeComponent();
          
            labName.Text = name;//名字
            label4.Text = team;//組別
            this.permission = permission;
            this.style = style;
            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);
        }

     
        

        private void Form2_Load(object sender, EventArgs e)  //具表單
        {//設定螢幕大小
            this.Size = new System.Drawing.Size(Screen.PrimaryScreen.WorkingArea.Width - 300, Screen.PrimaryScreen.WorkingArea.Height - 300);
            //主程式
            connString = "server=" + IP.ToString() + ";port=" + port.ToString() + ";user id="+labName.Text+";password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫
            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                                                       
            String selectused = @"select * from `taskrejected` ";      
            MySqlCommand cmdusd = new MySqlCommand(selectused, conn);

            MySqlDataReader drusd = cmdusd.ExecuteReader();
            while (drusd.Read())
            {              
                string TaskCard = drusd["TaskCard"].ToString();
                    string TaskCardCode = drusd["TaskCardCode"].ToString();
                string update = drusd["UpdateTime"].ToString();
                string[] u1 = update.ToString().Split('/');
                string Engineer = drusd["Engineer"].ToString();
                    string Manager = drusd["Manager"].ToString();
                    string Reason = drusd["Reason"].ToString();
                    string State = drusd["State"].ToString();
                
                string thisday = DateTime.Now.Year.ToString()+ "/" + DateTime.Now.Month.ToString();
              
                if (thisday.ToString().Contains(u1[0]) && thisday.ToString().Contains(u1[1]))
                    Gridsch.Rows.Add(new Object[] {TaskCard,TaskCardCode,update,Engineer,Manager,Reason,State });
                

            }
            drusd.Close();
            cmdusd.Dispose();
            conn.Close();

        }

    

        private void pictureBox4_Click(object sender, EventArgs e) //帳號動作按紐
        {
            AccountMenu.Show(pictureBox4, 0, pictureBox4.Height);
        }

        private void button7_Click(object sender, EventArgs e)//月份搜尋
        {
           /* if ((metroTextBox1.Text.Equals("1")) || (metroTextBox1.Text.Equals("2")) || (metroTextBox1.Text.Equals("3")) || (metroTextBox1.Text.Equals("4")) ||
                (metroTextBox1.Text.Equals("5")) || (metroTextBox1.Text.Equals("6")) || (metroTextBox1.Text.Equals("7"))|| (metroTextBox1.Text.Equals("8")) ||
                (metroTextBox1.Text.Equals("9")) || (metroTextBox1.Text.Equals("10")) || (metroTextBox1.Text.Equals("11")) || (metroTextBox1.Text.Equals("12")))*/
                if(true)
            {
                Gridsch.Rows.Clear();
       
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                String selectused = @"select * from `taskrejected`";
                MySqlCommand cmdusd = new MySqlCommand(selectused, conn);
                cmdusd.Parameters.Add(new MySqlParameter("@user", metroTextBox1.Text));
                MySqlDataReader drusd = cmdusd.ExecuteReader();
                while (drusd.Read())
                {
                    string TaskCard = drusd["TaskCard"].ToString();
                    string TaskCardCode = drusd["TaskCardCode"].ToString();
                    string update = drusd["UpdateTime"].ToString();
                    string[] u1 = update.ToString().Split('/');
                    string searchtime = u1[0] + "/" + u1[1];
                    string Engineer = drusd["Engineer"].ToString();
                    string Manager = drusd["Manager"].ToString();
                    string Reason = drusd["Reason"].ToString();
                    string State = drusd["State"].ToString();
                    Console.WriteLine(searchtime.ToString());
                    if (u1[1].ToString().Equals(metroTextBox1.Text) || u1[0].ToString().Equals(metroTextBox1.Text)||searchtime.ToString().Equals(metroTextBox1.Text))
                    Gridsch.Rows.Add(new Object[] { TaskCard, TaskCardCode, update, Engineer, Manager, Reason, State });


                }
                drusd.Close();
                cmdusd.Dispose();
                conn.Close();
                metroTextBox1.Text = "";
            }
            else {
                MetroFramework.MetroMessageBox.Show(this, "輸入資料不符合格是", "MESSAGE BOX", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
           

        }

        private void metroPanel1_Click(object sender, EventArgs e)//帳號動作按紐
        {
            //帳號動作按紐
            AccountMenu.Show(pictureBox4, 0, pictureBox4.Height);
        }

        private void 關閉ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void 帳號登出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
            Form1 f1 = new Form1();
            f1.IP = this.IP;
            f1.port = this.port;
            f1.style = this.style;
            f1.Visible = true;
        }

        private void 變更密碼ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Changeaccount change = new Changeaccount();
            change.style = this.style;
            change.IP = this.IP;
            change.port = this.port;
            change.Name = labName.Text;
            change.ShowDialog();
        }
        private void Gridsch_CellContentClick(object sender, DataGridViewCellEventArgs e)//退文表單追朔
        {
           Reject r1 = new Reject();
            r1.IP = this.IP;
            r1.port = this.port;
            r1.style = this.style;
            r1.taskname = Gridsch.Rows[e.RowIndex].Cells[0].Value.ToString();
            r1.Name = labName.Text;
             r1.ShowDialog();
        }

        private void textimportTask_Click(object sender, EventArgs e)//開啟Task按鈕
        {
        
            using (OpenFileDialog oOpenFileDialog = new OpenFileDialog())
            {
                oOpenFileDialog.Filter = " Word|*.doc| Word|*.docx| All Files|*.*";
                oOpenFileDialog.Title = "Select a  File";
                oOpenFileDialog.FilterIndex = 3;


                if (oOpenFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Alltaskgrid.Rows.Clear();
                    path = oOpenFileDialog.FileName;//絕對路徑
                    textimportTask.Text = path;
                    string fileName = Path.GetFileName(path);//設定路徑跟檔名 資料庫用
                                                             //工卡總表
               
                    MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
                    if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                    String tasksql = @"select TaskCardCode,TaskCardName,TaskLocation,TaskDetail,Equipment,Maintaincycle,Safe from `taskcard` ";
                    MySqlCommand taskcmd = new MySqlCommand(tasksql, conn);

                    MySqlDataReader taskdr = taskcmd.ExecuteReader();
                    while (taskdr.Read())
                    {

                        string code = taskdr["TaskCardCode"].ToString();
                        string Cardname = taskdr["TaskCardName"].ToString();
                        string location = taskdr["TaskLocation"].ToString();
                        string Detail = taskdr["TaskDetail"].ToString();
                        string equipment = taskdr["Equipment"].ToString();
                        string Maintain = taskdr["Maintaincycle"].ToString();
                        string Safe = taskdr["Safe"].ToString();
                        Alltaskgrid.Rows.Add(new Object[] { code, Cardname, location, Detail, equipment, Maintain, Safe });


                    }
                    taskdr.Close();
                    taskcmd.Dispose();
                    conn.Dispose();

                }
            }

           


        }

    
    
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var backgroundWorker = sender as BackgroundWorker;
            for (int j = 0; j < 200; j++)
            {
       
                backgroundWorker.ReportProgress(j);
            }
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pgbShow.Value = e.ProgressPercentage;
            //當backgroundWorker的i改變時就會觸發，進而更改pgbShow.Value
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // MessageBox.Show("Processing was completed");
            //當backgroundWorker工作完成時顯示
            pgbShow.Visible = false;
        }

        private void TaskBut_Click(object sender, EventArgs e)//工卡總表存取
        {//progressBard
            foreach (Process item in Process.GetProcessesByName("WINWORD"))
            { item.Kill(); }//如存在開啟的Word則先關閉(一個應用程式例項就是程序)
            backgroundWorker.WorkerReportsProgress = true;//啟動回報進度
            pgbShow.Maximum = 200;
            pgbShow.Step = 1;
            pgbShow.Value = 0;
            pgbShow.Visible = true;
            backgroundWorker.RunWorkerAsync();


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
                            pgbShow.Value++;                   
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
                          /*      case "工具 / 治具 / 機具設備需求 Tool / Jig / Equipment Requirement":
                                          need = 3;
                                    rowsql = rowPos + 2;
                                    deviceconstant++;
                                    break;*/
                                case "工具/治具/機具設備需求 Tool/Jig/Equipment Requirement":
                                    need = 3;
                                    rowsql = rowPos + 2;
                                    deviceconstant++;
                                    break;
                                default:
                                    break;
                                }
                            if(need==1 & rowsql<=rowPos)
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
                                    //  MessageBox.Show(people);
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
                                  
                                    // MessageBox.Show(material);
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
                        catch { Good = "";
                         //   MessageBox.Show(rowPos.ToString() + columPos + "test one");
                        }
                    

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
                    Alltaskcode = nowTable.Cell(1+ testadd, 1).Range.Text;
                    Alltaskcode = Alltaskcode.Remove(Alltaskcode.Length - 2, 2);//remove \r\a
                    Alltaskcode = Alltaskcode.Replace("工作說明書編碼", "").Replace("：", "");
                    Alltaskcode = Alltaskcode.Replace("Task Card Code", "");
                    Alltaskcode = Alltaskcode.TrimStart();
                   // Alltaskcode = splite(Alltaskcode);
                    Console.WriteLine(Alltaskcode.ToString());
                    Alltaskname = nowTable.Cell(1 + testadd, 2).Range.Text;
                    Alltaskname = Alltaskname.Remove(Alltaskname.Length - 2, 2);//remove \r\a\
                    Alltaskname = Alltaskname.Replace("工作說明書名稱", "").Replace("：","").Replace("◎SC", "");
                    Alltaskname = Alltaskname.Replace("Task Card Name", "");
                    Alltaskname = Alltaskname.TrimStart();
                    Console.WriteLine(Alltaskname.ToString());
                    Alltaskdetail = nowTable.Cell(2 + testadd, 1).Range.Text;
                    Alltaskdetail = Alltaskdetail.Remove(Alltaskdetail.Length - 2, 2);//remove \r\a
                    Alltaskdetail = Alltaskdetail.Replace("工作說明書簡述", "").Replace("：", "");
                    Alltaskdetail = Alltaskdetail.Replace("Task Card Description", "");
                    Alltaskdetail = Alltaskdetail.TrimStart();
                    Console.WriteLine(Alltaskdetail.ToString()); 


                    Alltasklocation = nowTable.Cell(3 + testadd, 1).Range.Text;
                    Alltasklocation = Alltasklocation.Remove(Alltasklocation.Length - 2, 2);//remove \r\a
                    Alltasklocation = Alltasklocation.Replace("維修適用範圍", "").Replace("：", "");
                    Alltasklocation = Alltasklocation.Replace("Location", "");
                    Alltasklocation = Alltasklocation.TrimStart();
                    Console.WriteLine(Alltasklocation.ToString());
                    Alltaskdevice = nowTable.Cell(3 + testadd, 2).Range.Text;
                    Alltaskdevice = Alltaskdevice.Remove(Alltaskdevice.Length - 2, 2);//remove \r\a
                    Alltaskdevice = Alltaskdevice.Replace("維修適用設備","").Replace("：", "");
                    Alltaskdevice = Alltaskdevice.Replace("Equipment Description", "");
                    Alltaskdevice = Alltaskdevice.TrimStart();
                    Console.WriteLine(Alltaskdevice.ToString());

                    maintain = nowTable.Cell(3 + testadd, 3).Range.Text;
                    maintain = maintain.Remove(maintain.Length - 2, 2);//remove \r\a
                    maintain = maintain.Replace("維修週期", "").Replace("：", "");
                    maintain = maintain.Replace("Task Interval", "").Replace("Interval","");
                    maintain = maintain.TrimStart();
                    Console.WriteLine(maintain.ToString());
                    safe = nowTable.Cell(4 + testadd, 1).Range.Text;
                    safe = safe.Remove(safe.Length - 2, 2);//remove \r\a
                    Console.WriteLine("first = "+safe.ToString());
                   
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
               /* if (deviceconstant == 0)
                {
                        nowTable = doc.Tables[2];
                        for (int rowPos = 1; rowPos <= nowTable.Rows.Count; rowPos++)//抓第幾----
                        {
                            for (int columPos = 1; columPos <= nowTable.Columns.Count; columPos++)//抓底幾個|||
                            {
                            try
                            {
                                Good = nowTable.Cell(rowPos, columPos).Range.Text;
                                Good = Good.Remove(Good.Length - 2, 2);
                            }
                            catch { Good = ""; }
                                switch (Good.ToString())
                                {
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
                                if (need == 3 & rowsql <= rowPos)
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
                                            default:
                                                MessageBox.Show("default");
                                                break;


                                        }

                                    }

                                }
                            }//colunn
                        }//row
                 
               }//the second table*/
                Console.WriteLine(people.ToString()+"\n"+ material.ToString()+"\n"+ Device.ToString());
                BuildFolderTask(Alltaskcode.ToString());
                ///寫進資料庫
                //   string lower = txtCode.Text;//轉換輸入格式
                //   lower = lower.ToLower().ToString();

           
              MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                try//測試有沒有表格
                {
                    string showTable = "desc `" + Alltaskcode.ToString().ToLower() + @"`";
                    MySqlCommand Tablefind = new MySqlCommand(showTable, conn);

                    //index1 = Tablefind.ExecuteNonQuery();
                    Tablefind.ExecuteNonQuery();
                    Console.WriteLine("進入資料庫有表格");
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
                        if (index > 0)
                            MessageBox.Show("update成功");
                        else
                            MessageBox.Show("update失敗");

                    }
                }

                catch //沒有表格
                {
                    Console.WriteLine("進入資料庫沒表格");
                    string Tablesql = @"CREATE TABLE `" + Alltaskcode.ToString().ToLower() + @"` (TaskCardName VARCHAR(45),Items int(11),TaskCardCode TEXT,TaskDetail TEXT,TaskImage TEXT,TaskImagepath TEXT,StandardValue VARCHAR(45),Note VARCHAR(45),Equipment SET('A','B'))";
                    MySqlCommand creatcmd = new MySqlCommand(Tablesql, conn);
                    int Create = creatcmd.ExecuteNonQuery();

               /*     if (Create > -1)
                        MessageBox.Show("創建TaskCard成功");
                    else
                        MessageBox.Show("創建TaskCard失敗");*/
                    ///////注意還未測試
                    string Checksql = @"CREATE TABLE `" + Alltaskcode.ToString().ToLower() + "-f1" + @"` (TaskCardName VARCHAR(45),Items int(11),TaskCardCode TEXT,TaskDetail TEXT,Equipment VARCHAR(45),StandardValue VARCHAR(45),Note VARCHAR(45),CheckResult VARCHAR(45) )";
                    MySqlCommand checkcreatecmd = new MySqlCommand(Checksql, conn);
                    int CheckCreate = checkcreatecmd.ExecuteNonQuery();

                   /* if (CheckCreate > -1)
                        MessageBox.Show("創建CheckList成功");
                    else
                        MessageBox.Show("創建CheckList失敗");*/

                    string Checkrecordsql = @"CREATE TABLE `" + Alltaskcode.ToString().ToLower() + "-f1-record" + @"` (Name VARCHAR(20),Items TEXT,TaskCardCode TEXT,CheckDetail TEXT,CheckResult TEXT,Result ENUM('Qualified','Unqualified',''),Time TIMESTAMP,FinishTime TEXT,TaskLocation ENUM('Zuoying_Station','Tainan_Station','Zuoying_Base','Yanchao_Factory'),Cycle TEXT,Sign VARCHAR(20),Location TEXT,GPS TEXT)";

                    MySqlCommand checkrecordcmd = new MySqlCommand(Checkrecordsql, conn);
                    int Checkrecord = checkrecordcmd.ExecuteNonQuery();

                  /*  if (Checkrecord > -1)
                        MessageBox.Show("創建Check-record成功");
                    else
                        MessageBox.Show("創建Check-record失敗");*/

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
                        if (index > 0)
                            MessageBox.Show("Import成功");
                        else
                            MessageBox.Show("Import失敗");
                    }
                }
                finally {
                    doc.Close();
                    textimportTask.Text = "";
                }
             
            
            
                
                
            }
            Alltaskgrid.Rows.Clear();
          
             MySqlConnection conn1 = new MySqlConnection(connString);//實做一個物件
            if (conn1.State != ConnectionState.Open) conn1.Open();//連線器打開
            String tasksql = @"select TaskCardCode,TaskCardName,TaskLocation,TaskDetail,Equipment,Maintaincycle,Safe from `taskcard` ";
            MySqlCommand taskcmd = new MySqlCommand(tasksql, conn1);

            MySqlDataReader taskdr = taskcmd.ExecuteReader();
            while (taskdr.Read())
            {

                string code = taskdr["TaskCardCode"].ToString();
                string Cardname = taskdr["TaskCardName"].ToString();
                string location = taskdr["TaskLocation"].ToString();
                string Detail = taskdr["TaskDetail"].ToString();
                string equipment = taskdr["Equipment"].ToString();
                string Maintain = taskdr["Maintaincycle"].ToString();
                string Safe = taskdr["Safe"].ToString();
                Alltaskgrid.Rows.Add(new Object[] { code, Cardname, location, Detail, equipment, Maintain, Safe });


            }
            taskdr.Close();
            taskcmd.Dispose();
            conn1.Dispose();
        }

        private string splite(string splite)
        {
            string re = "";
            char[] delimiterChars = { '\n', '\r', '\a' };//remove \r\a
            re = splite.ToString();
            string[] tests = re.Split(delimiterChars);
            if (tests[1].ToString().Equals("Task Card Code") || tests[1].ToString().Equals("Task Card Name")|| tests[1].ToString().Equals("Location") ||tests[1].ToString().Equals("Equipment Description")
                || tests[1].ToString().Equals("Task Interval")|| tests[1].ToString().Equals("Safety Precautions") || tests[1].ToString().Equals("Task Card Long Description")
                || tests[1].ToString().Equals("Task Card Code ") || tests[1].ToString().Equals("Task Card Name ") || tests[1].ToString().Equals("Location ID  ") || tests[1].ToString().Equals("Task Interval ")|| tests[1].ToString().Equals("Work Asset ID  "))//TSAVI-SL01-12_V9 信號線路12月檢
            { re = tests[2].ToString(); }

            else if (tests[1].ToString() != "")
                re = tests[1].ToString();


            return re.ToString();
        }

        ////////////////
        private void BuildFolderTask(string BF)//創建任務資料夾
        {
            string username = "test1";
            string password = "1234";
            string URLAddress = "ftp://163.18.57.239/";

            FtpWebRequest re = (FtpWebRequest)WebRequest.Create(URLAddress + BF.ToString().ToLower());
            re.Credentials = new NetworkCredential(username, password);
            re.Method = WebRequestMethods.Ftp.MakeDirectory;
            re.Timeout = (60000 * 1);
            try
            {
                FtpWebResponse response = (FtpWebResponse)re.GetResponse();
                response.Close();
            }
            catch { Console.WriteLine("此工卡總表有重複輸入"); }



        }

        private void schpath_Click(object sender, EventArgs e)//排程匯入PATH
        {
          
        }
        private void metroTile1_Click(object sender, EventArgs e)//排程匯入PATH
        {
            using (OpenFileDialog oOpenFileDialog = new OpenFileDialog())
            {
                oOpenFileDialog.Filter = " EXCEL|*.xls| Word|*.xlms| All Files|*.*";
                oOpenFileDialog.Title = "Select a  File";
                oOpenFileDialog.FilterIndex = 3;


                if (oOpenFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    excelpath = oOpenFileDialog.FileName;//絕對路徑
                    schpath.Text = excelpath;
                    excelname = Path.GetFileName(excelpath);//設定路徑跟檔名 資料庫用

                }
            }

        }
        public static string GetFirstSheetNameFromExcelFileName(string filepath, int numberSheetID)//搜尋EXCEL SHEETName
        {
            if (!System.IO.File.Exists(filepath))
            {
                return "This file is on the sky??";
            }
            if (numberSheetID <= 1) { numberSheetID = 1; }

            Microsoft.Office.Interop.Excel.Application obj = default(Microsoft.Office.Interop.Excel.Application);
            Microsoft.Office.Interop.Excel.Workbook objWB = default(Microsoft.Office.Interop.Excel.Workbook);
            string strFirstSheetName = null;

            obj = (Microsoft.Office.Interop.Excel.Application)Microsoft.VisualBasic.Interaction.CreateObject("Excel.Application", string.Empty);
            objWB = obj.Workbooks.Open(filepath, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            strFirstSheetName = ((Microsoft.Office.Interop.Excel.Worksheet)objWB.Worksheets[1]).Name;

            objWB.Close(Type.Missing, Type.Missing, Type.Missing);
            objWB = null;
            obj.Quit();
            obj = null;
            return strFirstSheetName;


        }



        private void schbut_Click(object sender, EventArgs e)//排程匯入
        {
            // Console.WriteLine(GetFirstSheetNameFromExcelFileName(excelpath, 1));
            // string sheetName = schloc.Text;
            if (!schpath.Text.ToString().Equals(""))
            {
                string sheetName = GetFirstSheetNameFromExcelFileName(excelpath, 1);

                string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + schpath.Text + "';Extended Properties='EXCEL 8.0;HDR=YES;IMEX=1;'"; //此連接只能操作Excel2007之前(.xls)文件
                System.Diagnostics.Process.Start(schpath.Text);

                using (OleDbConnection conn_excel = new OleDbConnection(strConn))


                {//Excel 8.0 是 Office 97 的 Excel 格式，這個格式從 Excel 97 ~ Excel 2003 都相容，如果你在 Excel 中使用「另存新檔」的話，必須要選取這個檔案類型儲存，才能正確用 OleDb 正確開
                 //若指定值為 Yes，代表 Excel 檔中的工作表第一列是欄位名稱            若指定值為 No，代表 Excel 檔中的工作表第一列就是資料了，沒有欄位名稱
                 //當 IMEX=0 時為「匯出模式」，這個模式開啟的 Excel 檔案只能用來做「寫入」用途。
                 //當 IMEX = 1 時為「匯入模式」，這個模式開啟的 Excel 檔案只能用來做「讀取」用途。
                 //當 IMEX = 2 時為「連結模式」，這個模式開啟的 Excel 檔案可同時支援「讀取」與「寫入」

                    conn_excel.Open();

                    OleDbCommand cmd_excel = new OleDbCommand("SELECT * FROM [" + sheetName + "$];", conn_excel);

                    OleDbDataReader reader_excel = cmd_excel.ExecuteReader();
                    //SQL連線字串
                
                    using (MySqlConnection cn_sql = new MySqlConnection(connString))

                    {

                        cn_sql.Open();

                        //宣告Transaction

                      //  MySqlTransaction stran = cn_sql.BeginTransaction();

                        try

                        {

                            while (reader_excel.Read())//列

                            {
                                string constant = "";
                                for (int i = 1; i <= 31; i++)
                                {
                                    //  MessageBox.Show(reader_excel[i + 4].ToString());
                                    if (reader_excel[i + 4].ToString() == "V" || reader_excel[i + 4].ToString() == "")
                                    {
                                        constant = i.ToString();
                                        // MessageBox.Show(constant);
                                    }


                                }

                                string Date = schyear.Text + "-" + reader_excel[1] + "-" + constant.ToString();

                                if (reader_excel[1].ToString().Equals("1") || reader_excel[1].ToString().Equals("2") || reader_excel[1].ToString().Equals("3") || reader_excel[1].ToString().Equals("4")//月份判斷
                                    || reader_excel[1].ToString().Equals("5") || reader_excel[1].ToString().Equals("6") || reader_excel[1].ToString().Equals("7") || reader_excel[1].ToString().Equals("8")
                                    || reader_excel[1].ToString().Equals("9") || reader_excel[1].ToString().Equals("10") || reader_excel[1].ToString().Equals("11") || reader_excel[1].ToString().Equals("12"))

                                {
                                    string t = "Task";

                                    string Locat = "";
                                    string go = reader_excel[3].ToString();

                                    bool SearchResult = go.Contains("左營車站");
                                    if (SearchResult.ToString().Equals("True"))
                                        Locat = "Zuoying_Station";
                                    bool SearchResult1 = go.Contains("台南車站");
                                    if (SearchResult1.ToString().Equals("True"))
                                        Locat = "Tainan_Station";
                                    bool SearchResult2 = go.Contains("左營基地");
                                    if (SearchResult2.ToString().Equals("True"))
                                        Locat = "Zuoying_Base";
                                    bool SearchResult3 = go.Contains("總機廠");
                                    if (SearchResult3.ToString().Equals("True"))
                                        Locat = "Yanchao_Factory";





                                    MySqlCommand cmd_sql = new MySqlCommand("insert into schedule_info (Items,Year,Month,TaskCardCode,Location_content,TaskDetail,Date,State,Location) values ('" + reader_excel[0] + "','" + schyear.Text + "','" + reader_excel[1] + "','" + reader_excel[2] + "','" + reader_excel[3] + "','" + reader_excel[4] + "','" + Date.ToString() + "','" + t.ToString() + "','" + Locat.ToString() + "')", cn_sql);

                                   

                                     cmd_sql.ExecuteNonQuery();
                                    
                                }


                            }
                            MessageBox.Show("Import成功");
                            //迴圈跑完並一次Insert

                            //stran.Commit();

                        }

                        catch (MySqlException ex)

                        {

                            MessageBox.Show(ex.Message);


                            //stran.Rollback();

                        }

                        catch (OleDbException ex)

                        {

                            MessageBox.Show(ex.Message);

                           //stran.Rollback();

                        }

                        catch (Exception ex)

                        {

                            MessageBox.Show(ex.Message);

                         //   stran.Rollback();

                        }

                        finally

                        {

                            cn_sql.Close();

                            conn_excel.Close();

                            reader_excel.Close();

                        }
                    }
                }
            }
          
        }

     
       
        private void TaskCheckIn_Click(object sender, EventArgs e)//進入標單匯入
        {
            TaskCheckWord task = new TaskCheckWord(labName.Text, label4.Text, this.permission);
            task.IP = this.IP;
            task.port = this.port;
            task.style = this.style;
            task.Name = labName.Text;
            task.ShowDialog();

        }
        private void metroButton1_Click(object sender, EventArgs e)//帳號權限 後台端
        {
            //後台管理端
            i2 = 1;//設置等等按工程師可以出現
            int i3 = 0;//選擇哪個ROW


            string permission = "";
            if (i == 1)//防止從副輸入到dataGridView1
            {
                metroGrid2.Rows.Clear();
         
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

                String selectGroup = @"select `Name`,`Team`,`Permission`,`Lasttime` from `account` where `Position` = @user";
                MySqlCommand cmd1 = new MySqlCommand(selectGroup, conn);
                string position = "backmanager";
                cmd1.Parameters.Add(new MySqlParameter("@user", position));
                MySqlDataReader dr1 = cmd1.ExecuteReader();
                while (dr1.Read())
                {
                    //放進DataGridView

                    string name = dr1["Name"].ToString();
                    string team = dr1["Team"].ToString();
                    permission = dr1["Permission"].ToString();
                    string time = dr1["Lasttime"].ToString();
                    if(permission !="5")
                    metroGrid2.Rows.Add(new Object[] { name, permission, team, position,time });
                    //Per(permission, i3);
                    //i3 = i3 + 1;

                }
                dr1.Close();
                conn.Close();
                cmd1.Dispose();

            }
            i++;//i=2代表data在後台
            //i2=2代表在工程師
        }

        private void metroButton2_Click(object sender, EventArgs e)//帳號權限 手機端
        {//工程師
            i = 1;
            int i3 = 0;
            string permission = "";

            if (i2 == 1)
            {
                metroGrid2.Rows.Clear();
             
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

                String selectGroup = @"select `Lasttime`,`Name`,`Team`,`Permission` from `account` where `Position` = @user";
                MySqlCommand cmd1 = new MySqlCommand(selectGroup, conn);
                string position = "engineer";
                cmd1.Parameters.Add(new MySqlParameter("@user", position));
                MySqlDataReader dr1 = cmd1.ExecuteReader();
                while (dr1.Read())
                {
                    //放進DataGridView

                    string name = dr1["Name"].ToString();
                    string team = dr1["Team"].ToString();
                    permission = dr1["Permission"].ToString();
                    string time = dr1["Lasttime"].ToString();
                    if (permission != "5")
                        metroGrid2.Rows.Add(new Object[] { name, permission, team, position,time });
                    //Per(permission, i3);
                    //i3 = i3 + 1;
                }
                dr1.Close();
                conn.Close();
                cmd1.Dispose();

            }
            i2++;

        }
        private void metroButton4_Click(object sender, EventArgs e)//督導端
        {
            metroGrid2.Rows.Clear();
           
            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

            String selectGroup = @"select `Lasttime`,`Name`,`Team`,`Permission` from `account` where `Position` = @user";
            MySqlCommand cmd1 = new MySqlCommand(selectGroup, conn);
            string position = "manager";
            cmd1.Parameters.Add(new MySqlParameter("@user", position));
            MySqlDataReader dr1 = cmd1.ExecuteReader();
            while (dr1.Read())
            {
                //放進DataGridView

                string name = dr1["Name"].ToString();
                string team = dr1["Team"].ToString();
                string managerper = dr1["Permission"].ToString();
                string time = dr1["Lasttime"].ToString();
                if (managerper != "5")
                    metroGrid2.Rows.Add(new Object[] { name, managerper, team, position, time });
                //Per(permission, i3);
                //i3 = i3 + 1;
            }
            dr1.Close();
            conn.Close();
            cmd1.Dispose();

        }

        private void metroButton3_Click(object sender, EventArgs e)//帳號權限存取
        {
            try
            {
                foreach (DataGridViewRow row in metroGrid2.Rows)
                {
                    string Column1 = row.Cells[0].Value.ToString();//Name

                    string Column2 = row.Cells[1].Value.ToString();//permission


                    if (Column2 == "")
                    {
                        //   MessageBox.Show("No");
                    }
                    else
                    {
                        MySqlConnection connone = new MySqlConnection(connString);//實做一個物件來連線
                        if (connone.State != ConnectionState.Open) connone.Open();//連線器打開
                        string sql = @"UPDATE  account set Permission =@test2 ,Used = 1 where Name =@test1";
                        using (MySqlCommand cmd = new MySqlCommand(sql, connone))//匯入檔案到資料庫
                        {
                            cmd.Parameters.Add("@test1", Column1.ToString());
                            cmd.Parameters.Add("@test2", Column2.ToString());

                            int index = cmd.ExecuteNonQuery();

                            cmd.Dispose();
                            connone.Close();
                        }

                    }
                
                    
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
            String selectGroupclose = @"REVOKE EXECUTE, PROCESS, SHOW DATABASES, SHOW VIEW, ALTER, ALTER ROUTINE, CREATE, CREATE ROUTINE, CREATE TABLESPACE, CREATE TEMPORARY TABLES, CREATE VIEW, DELETE, DROP, EVENT, INDEX, INSERT, REFERENCES, TRIGGER, UPDATE, CREATE USER, FILE, GRANT OPTION, LOCK TABLES, RELOAD, REPLICATION CLIENT, REPLICATION SLAVE, SHUTDOWN, SUPER ON *.* FROM '"+Column1.ToString()+"'@'%';FLUSH PRIVILEGES";
            MySqlCommand cmd1close = new MySqlCommand(selectGroupclose, conn);
           
           
            MySqlDataReader dr1close = cmd1close.ExecuteReader();
            MessageBox.Show("帳號以降低權限","警告",MessageBoxButtons.OK,MessageBoxIcon.Warning);

                    MySqlConnection conn1 = new MySqlConnection(connString);//實做一個物件
                        if (conn1.State != ConnectionState.Open) conn1.Open();//連線器打開

                        string selectGroup = "";
                        switch (Column2)
                        {
                            case "0":
                                selectGroup = "GRANT SELECT, SHOW DATABASES, SHOW VIEW, INSERT, UPDATE, DELETE ON *.* TO '" + Column1.ToString() + "'@'%'";
                            break;

                            case "1":
                                selectGroup = "GRANT SELECT, SHOW DATABASES, SHOW VIEW ON *.* TO '" + Column1.ToString() + "'@'%'";
                            break;

                            case "2":
                               selectGroup = "GRANT SELECT, SHOW DATABASES, SHOW VIEW, INSERT ON *.* TO '" + Column1.ToString()+"'@'%'";
                            break;

                            case "3":
                                selectGroup = "GRANT SELECT, SHOW DATABASES, SHOW VIEW, INSERT, UPDATE ON *.* TO '" + Column1.ToString() + "'@'%'";
                            break;

                            case "4":
                                selectGroup = "GRANT SELECT, SHOW DATABASES, SHOW VIEW, INSERT, UPDATE, DELETE ON *.* TO '" + Column1.ToString() + "'@'%'";
                            break;

                            case "5":
                                selectGroup = " GRANT EXECUTE, PROCESS, SELECT, SHOW DATABASES, SHOW VIEW, ALTER, ALTER ROUTINE, CREATE, CREATE ROUTINE, CREATE TABLESPACE, CREATE TEMPORARY TABLES, CREATE VIEW, DELETE, DROP, EVENT, INDEX, INSERT, REFERENCES, TRIGGER, UPDATE, CREATE USER, FILE, LOCK TABLES, RELOAD, REPLICATION CLIENT, REPLICATION SLAVE, SHUTDOWN, SUPER ON *.* TO '" + Column1.ToString() + "'@'%'";
                            break;
                        }
                    MessageBox.Show(selectGroup);
                        MySqlCommand cmd1 = new MySqlCommand(selectGroup, conn1);


                        MySqlDataReader dr1 = cmd1.ExecuteReader();
                  
                      
                        //MessageBox.Show("權限已調整", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        
                }



            }
            catch (Exception ex){ MessageBox.Show(ex.ToString()); }

        }

        private void metroGrid2_CellContentClick(object sender, DataGridViewCellEventArgs e)//判斷修改權限
        {//if (i == 1 || i2 == 1)
            {if (Int32.Parse(this.permission) == 5)//登入身分全縣
                {
                    int permissionper = 0;
                    try { permissionper = Int32.Parse(metroGrid2.Rows[e.RowIndex].Cells[1].Value.ToString()); }
                    catch { permissionper = -1; }
                    string permissionName = metroGrid2.Rows[e.RowIndex].Cells[0].Value.ToString();
                    string permissionTeam = metroGrid2.Rows[e.RowIndex].Cells[2].Value.ToString();
                    string permissionposition = metroGrid2.Rows[e.RowIndex].Cells[3].Value.ToString();
                    string permissionTime = metroGrid2.Rows[e.RowIndex].Cells[4].Value.ToString();
                    Changepermission p1 = new Changepermission();
                    p1.style = this.style;
                    p1.permiss = permissionper;
                    p1.name = permissionName;
                    p1.team = permissionTeam;
                    p1.position = permissionposition;
                    DialogResult dr = p1.ShowDialog();
                    if (dr == DialogResult.OK)
                    {

                        if (Int32.Parse(this.permission) > p1.Getper())
                        {
                            metroGrid2.Rows[e.RowIndex].Cells[1].Value = p1.Getper();
                            
                        }
                        
                        else
                            MetroFramework.MetroMessageBox.Show(this, "未到達可修改的權限值", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else if (dr == DialogResult.Cancel)
                    {
                        //MessageBox.Show(input.GetchanMsg());
                    }
                }
               
            }
           
        }

      
        private void CodeTitle_Click(object sender, EventArgs e)//地方權限解除
        {// comboBox1.Visible = true;
            
            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

            String selectused = @"select Region from `locationgps` ";
            MySqlCommand cmdusd = new MySqlCommand(selectused, conn);

            MySqlDataReader drusd = cmdusd.ExecuteReader();
            try
            {
                while (drusd.Read())
                {
                    var writer = new BarcodeWriter  //dll裡面可以看到屬性
                    {
                        Format = BarcodeFormat.QR_CODE,
                        Options = new QrCodeEncodingOptions //設定大小
                        {
                            Height = 300,
                            Width = 300,
                        }
                    };

                    pictureBox1.Image = writer.Write(drusd["Region"].ToString()); //轉QRcode的文字    
                    savephoto(drusd["Region"].ToString());
                    

                }
                MessageBox.Show("成功儲存");
            }
            catch { MessageBox.Show("路徑不存在"); }
            finally {  }

            // comboBox1.Visible = false;

        }
      
        private void SQLback_Click(object sender, EventArgs e)//資料庫的備份
        {
          
            string file = @"D:\SCD\SQLBackup\SQLbackup_" + DateTime.Now.ToString("MM_dd_yyyy") + ".sql";//資料庫備份本機端路徑
            using (MySqlConnection conn = new MySqlConnection(connString))//實做一個物件
            { if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                using (MySqlCommand cmd = new MySqlCommand())
                {
                    using (MySqlBackup mb = new MySqlBackup(cmd))
                    {
                        try
                        {
                            cmd.Connection = conn;
                             mb.ExportToFile(file);
                            conn.Close();
                            Console.WriteLine("SQLbackup    =  Done");
                            //傳爆FTP上
                            string username = "test1";
                            string password = "1234";
                            string uploadUrl = "ftp://163.18.57.239/" + "資料庫備份/" + DateTime.Now.ToString("MM_dd_yyyy") + ".sql";


                            if (!System.IO.File.Exists(@"D:\SCD\SQLBackup\SQLbackup_" + DateTime.Now.ToString("MM_dd_yyyy") + ".sql"))
                                MessageBox.Show("上傳檔案不存在");
                            // 要上傳的檔案               
                            else
                            {
                                string MyFileName = @"D:\SCD\SQLBackup\SQLbackup_" + DateTime.Now.ToString("MM_dd_yyyy") + ".sql";

                                WebClient wc = new WebClient();
                                wc.Credentials = new NetworkCredential(username, password);
                                wc.UploadFile(uploadUrl, MyFileName);
                                MessageBox.Show("成功 ， 儲存的路徑為=" + uploadUrl.ToString());
                            }


                        }
                        catch (Exception ex1) { MessageBox.Show(ex1.ToString()); }
                    }

                }
            }
        }
        private void FTPUpload(string photoname, string date)//地方權限
        {


            string username = "test1";
            string password = "1234";
            string uploadUrl = "ftp://163.18.57.239/" + "解開權限/" + date.ToString() + "/" + photoname.ToString();
            Console.WriteLine("儲存的路徑為=" + uploadUrl.ToString());

            if (!System.IO.File.Exists(@"D:\SCD\QRcode解鎖\" + photoname.ToString()))
                MessageBox.Show("上傳檔案不存在");

            // 要上傳的檔案
            //圖片格式
            else
            {
                string MyFileName = @"D:\SCD\QRcode解鎖\" + photoname.ToString();

                WebClient wc = new WebClient();
                wc.Credentials = new NetworkCredential(username, password);
                wc.UploadFile(uploadUrl, MyFileName);
                // path = uploadUrl.ToString();
                ///文件格式上傳 wc.UploadData(uploadUrl, data);
            }


        }
        private void savephoto(string Filename)//地方權限
        {
            //   Image image = Clipboard.GetImage(); 讀圖片給PictureBox
            //   pictureBox1.Image = image; 
            string fs = @"D:\SCD\QRcode解鎖\" + Filename.ToString() + ".jpg";


            Console.WriteLine("test for file name =" + fs.ToString());
            pictureBox1.Image.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);

            // MessageBox.Show(date.ToString());

            // BuildFolder(date.ToString());
            Console.WriteLine("file name=" + Filename.ToString());
            BuildFolder(Filename.ToString() + ".jpg");


        }
        private void BuildFolder(string BF)//地方權限
        {
            string date = DateTime.Now.ToShortDateString();
            date = date.Replace("/", "_");
            Console.WriteLine("date = " + date.ToString());
            string username = "test1";
            string password = "1234";
            string URLAddress = "ftp://163.18.57.239/" + "解開權限/";

            FtpWebRequest re = (FtpWebRequest)WebRequest.Create(URLAddress + date.ToString().ToLower());
            re.Credentials = new NetworkCredential(username, password);
            re.Method = WebRequestMethods.Ftp.MakeDirectory;
            re.Timeout = (60000 * 1);
            try
            {
                FtpWebResponse response = (FtpWebResponse)re.GetResponse();
                response.Close();
            }
            catch { }

            // MessageBox.Show("創建資料表");
            FTPUpload(BF.ToString(), date.ToString());

        }

        private void metroDateTime1_CloseUp(object sender, EventArgs e)//表單預檢日期搜尋
        {   //日期搜尋
            metroGrid3.Rows.Clear();
            string ADate = metroDateTime1.Text.Replace('年', '-').Replace('月', '-').TrimEnd('日');
            string Month = ADate.Substring(5, 2).TrimEnd('-');

            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                                                                //抓登入權限
 //  String selectused = @"select TaskCardCode,Location,Date,Items,TaskDetail,Location_content from `schedule_info` where Date = @user";
            String selectused = @"select * from `schedule_info` where Date = @user";
            MySqlCommand cmdusd = new MySqlCommand(selectused, conn);
            cmdusd.Parameters.Add(new MySqlParameter("@user", ADate));
            MySqlDataReader drusd = cmdusd.ExecuteReader();
            while (drusd.Read())
            {
                if (drusd["Items"].ToString() != "-1")
                {
                    string item = drusd["Items"].ToString();
                    string year = drusd["Year"].ToString();
                    string month = drusd["Month"].ToString();
                    string Date = drusd["Date"].ToString().Remove(10).TrimEnd('上');
                    string Code = drusd["TaskCardCode"].ToString();
                    string locat = drusd["Location_content"].ToString();
                    string Location = drusd["Location"].ToString();
                    string Detail = drusd["TaskDetail"].ToString();
                    string State = drusd["State"].ToString();
                    string team = drusd["Team"].ToString();
                    metroGrid3.Rows.Add(new Object[] { item, year, month, Date, Code, locat, Detail, Location, State, team });
                }

            }
            drusd.Close();
            cmdusd.Dispose();
            String selectmonth = @"select * from `schedule_info` where Month = @user";

            MySqlCommand cmd = new MySqlCommand(selectmonth, conn);
            cmd.Parameters.Add(new MySqlParameter("@user", Month));
            MySqlDataReader drmonth = cmd.ExecuteReader();
            ADate = ADate.Replace('-', '/');
            while (drmonth.Read())
            {
                if (drmonth["Items"].ToString() != "-1")
                {
                    string item = drmonth["Items"].ToString();
                    string year = drmonth["Year"].ToString();
                    string month = drmonth["Month"].ToString();
                    string Date = drmonth["Date"].ToString().Remove(10).TrimEnd('上');
                    string Code = drmonth["TaskCardCode"].ToString();
                    string locat = drmonth["Location_content"].ToString();
                    string Location = drmonth["Location"].ToString();
                    string Detail = drmonth["TaskDetail"].ToString();
                    string State = drmonth["State"].ToString();
                    string team = drmonth["Team"].ToString();
                   
                    if (!Date.ToString().Contains(ADate.ToString()))
                        metroGrid3.Rows.Add(new Object[] { item, year, month, Date, Code, locat, Detail, Location, State, team });
                  

                }

            }
            drmonth.Close();
            cmd.Dispose();
            conn.Close();
        }

        private void search_Click(object sender, EventArgs e)//Code搜尋
        {
            //TaskCardCode搜尋法
            if (!(Codsearch.Text.Equals("")))
            {
                metroGrid3.Rows.Clear();
            
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                                                                    //抓登入權限
                String selectused = @"select * from `schedule_info` ";
                MySqlCommand cmdusd = new MySqlCommand(selectused, conn);

                MySqlDataReader drusd = cmdusd.ExecuteReader();
                while (drusd.Read())
                {
                    if (drusd["Items"].ToString() != "-1")
                    {
                        string item = drusd["Items"].ToString();
                        string year = drusd["Year"].ToString();
                        string month = drusd["Month"].ToString();
                        string Date = drusd["Date"].ToString().Remove(10).TrimEnd('上');
                        string Code = drusd["TaskCardCode"].ToString();
                        string locat = drusd["Location_content"].ToString();
                        string Location = drusd["Location"].ToString();
                        string Detail = drusd["TaskDetail"].ToString();
                        string State = drusd["State"].ToString();
                        string team = drusd["Team"].ToString();
                        if (Detail.ToString().Contains(Codsearch.Text))

                            metroGrid3.Rows.Add(new Object[] { item, year, month, Date, Code, locat, Detail, Location, State, team });

                     
                          
                    }

                }
                drusd.Close();
                cmdusd.Dispose();
                conn.Close();
            }
            Codsearch.Text = "";
        }


        private void metroGrid3_CellContentClick(object sender, DataGridViewCellEventArgs e)//點選進入表單觀看
        {
            string nameIn_Com = "";
            string Code = "";
            string Name = "";
            string location = "";
            //跳脫視窗
            try
            {
                 Code = metroGrid3.Rows[e.RowIndex].Cells[4].Value.ToString();
                 Name = metroGrid3.Rows[e.RowIndex].Cells[6].Value.ToString();
                 location = metroGrid3.Rows[e.RowIndex].Cells[7].Value.ToString();
               
            }
            catch { }
            

            Name = sp(Name.ToString());
            //偵測有無此表格
  
            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
            try//測試有沒有表格
            {
                string showTable = "desc `" + Name.ToString().ToLower() + @"`";
                MySqlCommand Tablefind = new MySqlCommand(showTable, conn);

                //index1 = Tablefind.ExecuteNonQuery();
                Tablefind.ExecuteNonQuery();
                ReadTask r1 = new ReadTask();//產生Form2的物件，才可以使用它所提供的Method
                r1.IP = this.IP;
                r1.port = this.port;
                r1.Code = Name.ToString();
                r1.style = this.style;
                r1.location = location.ToString();
                r1.TaskCodeonly = Code.ToString();
                r1.Name = labName.Text;
                r1.accountname =labName.Text;
                if (r1.ShowDialog() == DialogResult.OK)
                {
                    MessageBox.Show(r1.ToString());
                }

            }
            catch //沒有表格
            {
                MetroFramework.MetroMessageBox.Show(this, "還尚未匯入資料", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


     

        private void schyear_Click(object sender, EventArgs e)//快捷鍵設置
        {
            this.AcceptButton = schbut;
        }

        private void Codsearch_Click(object sender, EventArgs e)//快捷鍵設置
        {
            this.AcceptButton = search;
        }
        private string sp(string splite)
        {
            string re = "";
            char[] delimiterChars = { '\n', '\t', '\r', '\a', '_' };//remove \r\a
            re = splite.ToString();
            string[] tests = re.Split(delimiterChars);
            if (tests[0].ToString() != "")
                re = tests[0].ToString();
             return re.ToString();
        }

        private void lockToolStripMenuItem_Click(object sender, EventArgs e)//解鎖畫面
        {
            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開

            String selectGroup = @"select * from `account` where Name = @user";
            MySqlCommand cmd1 = new MySqlCommand(selectGroup, conn);
             cmd1.Parameters.Add(new MySqlParameter("@user", labName.Text));
            MySqlDataReader dr1 = cmd1.ExecuteReader();
            while (dr1.Read())
            {
                Lock r1 = new Lock();//產生Form2的物件，才可以使用它所提供的Method
                r1.Password = dr1["Password"].ToString();
                r1.Name = labName.Text;
                //this.Close();
                r1.ShowDialog();
                
            }
            dr1.Close();
            conn.Close();
            cmd1.Dispose();
         
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
         
        }

        private void metroTextBox1_Click(object sender, EventArgs e)//快捷鍵設置
        {
            this.AcceptButton = button7;
        }

        private void pictureBox5_Click(object sender, EventArgs e)//設計者
        {

        }

        private void 稽核模式ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
            String selectGroup = @"REVOKE EXECUTE, PROCESS, SHOW DATABASES, SHOW VIEW, ALTER, ALTER ROUTINE, CREATE, CREATE ROUTINE, CREATE TABLESPACE, CREATE TEMPORARY TABLES, CREATE VIEW, DELETE, DROP, EVENT, INDEX, INSERT, REFERENCES, TRIGGER, UPDATE, CREATE USER, FILE, GRANT OPTION, LOCK TABLES, RELOAD, REPLICATION CLIENT, REPLICATION SLAVE, SHUTDOWN, SUPER ON *.* FROM '"+labName.Text+"'@'%';FLUSH PRIVILEGES";
            MySqlCommand cmd1 = new MySqlCommand(selectGroup, conn);
           
           
            MySqlDataReader dr1 = cmd1.ExecuteReader();
            MessageBox.Show("帳號以降低權限","警告",MessageBoxButtons.OK,MessageBoxIcon.Warning);

        }

      

        private void pictureBox4_MouseUp(object sender, MouseEventArgs e)
        {

         //   pictureBox4.Image = Resource1.user1;
        }

        private void pictureBox4_MouseLeave(object sender, EventArgs e)
        {
            pictureBox4.Image = Resource1.user;
        }

        private void pictureBox4_MouseEnter(object sender, EventArgs e)
        {
            pictureBox4.Image = Resource1.user1;
        }

        private void TaskCardlab_Click(object sender, EventArgs e)
        {

        }

        private void metroButton5_Click(object sender, EventArgs e)//廣播存取
        {if (!(metroTextBox2.Text == "" || metroTextBox3.Text == ""))
            {
                MySqlConnection conn = new MySqlConnection(connString);//實做一個物件
                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                String selectused = @"select * from `account` where Name =@user";
                MySqlCommand cmdusd = new MySqlCommand(selectused, conn);
                cmdusd.Parameters.Add(new MySqlParameter("@user", labName.Text));
                MySqlDataReader drusd = cmdusd.ExecuteReader();
                string team = "";
                while (drusd.Read())
                {
                    team = drusd["Team"].ToString();
                }
                cmdusd.Dispose(); drusd.Dispose();

                string sql = @"INSERT INTO `" + "broadcast" + @"`(`Name`,`Team`,`Title`,`Innertext`,`StartTime`,`EndTime`) VALUES
                           (@test1,@test2,@test3,@test4,@test5,@test6) ";
                using (MySqlCommand cmd = new MySqlCommand(sql, conn))//匯入檔案到資料庫
                {

                    cmd.Parameters.Add("@test1", labName.Text);
                    cmd.Parameters.Add("@test2", team.ToString());
                    cmd.Parameters.Add("@test3", metroTextBox2.Text);
                    cmd.Parameters.Add("@test4", metroTextBox3.Text);
                    cmd.Parameters.Add("@test5", dateTimePicker1.Text);
                    cmd.Parameters.Add("@test6", dateTimePicker2.Text);


                    int index = cmd.ExecuteNonQuery();
                    bool success = false;
                    if (index > 0)
                        MessageBox.Show("廣播設置成功");
                    metroTextBox2.Text = ""; metroTextBox3.Text = "";
                }
            }
            else
            { MetroFramework.MetroMessageBox.Show(this, "資料尚未填寫完畢", "MESSAGE BOX", MessageBoxButtons.OK, MessageBoxIcon.Warning); metroTextBox2.Text = ""; metroTextBox3.Text = ""; }
            
        }

        private void metroTextBox2_Click(object sender, EventArgs e)
        {
            this.AcceptButton = metroButton5;
        }

        private void Tile_account_manager_Click(object sender, EventArgs e)
        {
            string manager_acoount_text = "";
            using (OpenFileDialog oOpenFileDialog = new OpenFileDialog())
            {
                oOpenFileDialog.Filter = " All Files|*.*";
                oOpenFileDialog.Title = "Select a  File";
                oOpenFileDialog.FilterIndex = 3;


                if (oOpenFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    
                    manager_acoount_text = oOpenFileDialog.FileName;//絕對路徑
                    account_manage(manager_acoount_text);
                    /*****************開個視窗解釋******/
                }
            }
        }

        private void account_manage(string text)
        {
            int row = 0;
            int column_name = 0;
            int coilmn_permission = 0;
            /*****************讀檔**************/

            /**********************************/
            /*while (text.Read())
             { row++;
                if(row == 1)
                {
                 for (int i=0;i< 檔案欄位數;i++)
                   {
                    if(檔案[i] == "name")
                        column_name = i;
                    if(檔案[i] == "permission")
                        column_permission = i;
                   }
                }
                else 
                {檔案[column_name]存取到資料庫 & 檔案[column_permission]存取到 資料庫

                }
                  
             }
                
            */
        }
    }
}
