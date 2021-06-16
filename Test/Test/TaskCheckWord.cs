using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Net;
using System.Diagnostics;

namespace Test
{
    public partial class TaskCheckWord : MetroFramework.Forms.MetroForm
    {
        public string IP = "";
        public string port = "";
        public int style = 0;
        public string Name = "";
        string Good;
        string Item;
        string Detail;
        string Detail1 = "";//項目細節輸出
        string standard;
        string Note;
        int i = 0;//進去資料庫的圖檔名
        string photoname = "";
        string path = "";
        string fileName = "";//檔案名
        int index = 0;//判斷資料庫有沒有進去
        int count = 1;//圖片匯入的順序
        string[] words = { };//更改後的程式
        string lower = "";//TASKCARD
        string nametask = "";

        //帳號
        public string Username = "";
        public string thisteam = "";
        public string permission = "";


        //        資料庫痊癒宣告
       string connString="";

        
        public  TaskCheckWord(string name, string team, string per)
        {
            InitializeComponent();
            Username = name;
            thisteam = team;
            permission = per;
           

        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog oOpenFileDialog = new OpenFileDialog())
            {
                oOpenFileDialog.Filter = " Word|*.doc| Word|*.docx| All Files|*.*";
                oOpenFileDialog.Title = "Select a  File";
                oOpenFileDialog.FilterIndex = 3;


                if (oOpenFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    path = oOpenFileDialog.FileName;//絕對路徑
                    textBox1.Text = path;
                    fileName = Path.GetFileName(path);//設定路徑跟檔名 資料庫用
                }
            }
        }
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var backgroundWorker = sender as BackgroundWorker;
            for (int j = 0; j < 450; j++)
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
        private void button2_Click(object sender, EventArgs e)
        {
            pgbShow.Value = 0;
               count = 1;
            foreach (Process item in Process.GetProcessesByName("WINWORD"))
            { item.Kill(); }//如存在開啟的Word則先關閉(一個應用程式例項就是程序)
          connString = "server= "+IP.ToString()+";port= "+port.ToString()+";user id=thsr2019_05_08;password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫
           // connString = "server=" +"163.18.57.236" + ";port=" + "3306" + ";user id=thsr2019_05_08;password=CzqhTlz0erd13UX6;database=thsr v1;charset=utf8;";//連線資料庫
            MySqlConnection conn = new MySqlConnection(connString);//實做一個物件來連線
            if (conn.State != ConnectionState.Open) conn.Open();//連線器打開



          

            if (!textBox1.Text.ToString().Equals(""))
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

            ///抓WORD圖片
            /* foreach (Microsoft.Office.Interop.Word.InlineShape ish in doc.InlineShapes)
              {
                     if ((ish.Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapeLinkedPicture) || (ish.Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture))
                     {
                         ish.Select();
                         app.Selection.Copy();
                         Image image = Clipboard.GetImage();
                         pictureBox1.Image = image;
                         string fs = @"C:\Users\B510\Desktop\File\" + i.ToString() + ".jpg";
                         pictureBox1.Image.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);
                         FTPUpload();
                         //FileStream fs = new FileStream(@"C:\Users\B510\Desktop\File\" + i.ToString() +".jpg", FileMode.Create);
                         // pictureBox1.Image.Save(fs + @"C:\Users\B510\Desktop\File\" + i.ToString() + ".jpg");
                         //  Image imgSave = pictureBox1.Image;
                         //    imgSave.Save(@"C:\Users\B510\Desktop\File\"+i.ToString());
                         //  MessageBox.Show("圖片" + i + "上傳成功");
                         i++;

                     }
              }*/
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
                        if ((taskcardtest.ToString().Contains("工作說明書") && taskcardtest.ToString().Contains("TASK CARD"))||(taskcardtest.ToString().Contains("工作說明書") && taskcardtest.ToString().Contains("Task Card")&&!(taskcardtest.ToString().Contains("編碼")) ))
                        //  if (taskcardtest.ToString().Contains("工作說明書") && taskcardtest.ToString().Contains("TASK CARD"))
                        {
                            //  MessageBox.Show("rowPos+1");
                            testadd =1;
                        }
                        //  string tableMessage = "";

                        for (int rowPos = 1+ testadd; rowPos <= nowTable.Rows.Count; rowPos++)//抓第幾----
                    {
                        Detail1 = "";

                        for (int columPos = 1; columPos <= nowTable.Columns.Count; columPos++)//抓底幾個|||
                            {
                              
  pgbShow.Value++;
                                
                                try
                                { Good = nowTable.Cell(rowPos, columPos).Range.Text;
                                    Good = Good.Remove(Good.Length - 2, 2);//remove \r\a
                                    //MessageBox.Show("Rowpos = "+rowPos + "\ncolumPos = " + columPos +"\n"+ Good.ToString());
                                }catch { //MessageBox.Show(columPos.ToString()+"        "+rowPos);
                                }



                                    if (Good.ToString().Equals("N/A"))
                                    Good = "";
                                switch (columPos)
                                {
                                    case 1:
                                        Item = Good.TrimStart().TrimEnd();

                                        if (rowPos == 1+testadd)
                                        {//取TaskCardCode
                                           // MessageBox.Show("Rowpos = " + rowPos + "\ncolumPos = " + columPos + "\n" + Item.ToString());
                                            Item = Item.Replace("工作說明書編碼", "").Replace("：", "").Replace("Task Card Name", "");
                                            Item = Item.Replace("Task Card Code", "");
                                            Item = Item.TrimStart().TrimEnd();
                                          //  char[] delimiterChars = { '\n', '\t', '\r', '\a' };//remove \r\a
                                            lower = Item.ToString();
                                          //  MessageBox.Show("Rowpos = " + rowPos + "\ncolumPos = " + columPos + "\n" + lower.ToString());
                                            //    MessageBox.Show(lower.ToString());
                                            // string[] tests = lower.Split(delimiterChars);
                                            //System.Console.WriteLine($"{words.Length} words in text:");
                                            //  lower = tests[1].ToString().ToLower();

                                            string delete = @"DELETE FROM`" + lower.ToString() + "`";
                                            MySqlCommand delcmd = new MySqlCommand(delete, conn);
                                            index = delcmd.ExecuteNonQuery();
                                            delcmd.Dispose();
                                            //MessageBox.Show("刪除");
                                            string deletecheck = @"DELETE FROM`" + lower.ToString() + "-f1`";
                                            MySqlCommand delcmdcheck = new MySqlCommand(deletecheck, conn);
                                            index = delcmdcheck.ExecuteNonQuery();
                                            delcmdcheck.Dispose();
                                           //MessageBox.Show("刪除");
 backgroundWorker.WorkerReportsProgress = true;//啟動回報進度
                                            pgbShow.Maximum = 450;
                                            pgbShow.Step = 1;
                                            pgbShow.Value = 0;
                                            pgbShow.Visible = true;
                                            backgroundWorker.RunWorkerAsync();
                                        }
                                        break;
                                    case 2:
                                        Detail = Good;
                                        if (rowPos == 1 + testadd)
                                        {//取工作內容
                                            Detail = Detail.Replace("工作說明書名稱", "").Replace("：", "").Replace("◎SC", "");
                                            Detail = Detail.Replace("Task Card Name", "");
                                            Detail = Detail.TrimStart();
                                            //char[] delimiterChars = { '\n', '\t', '\r', '\a' };//remove \r\a
                                            nametask = Detail.ToString();//轉換輸入格式
                                          //  MessageBox.Show(nametask.ToString());
                                           // words = nametask.Split(delimiterChars);
                                           // nametask = words[1].ToString();

                                        }
                                        else
                                        {//資料整理
                                            char[] delimiterChars = {'\n', '\t', '\r', '\a' };//remove \r\a
                                            name = Detail.ToString();//轉換輸入格式
                                            string[] words = name.Split(delimiterChars);
                                            int photoint = 0;
                                            for (int x = 0; x < words.Length; x++)
                                            { Console.WriteLine(words[x].ToString()+"長度為"+words.Length+"現在執行到"+x);
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

                                                else if (words[x].ToString().Equals("//") || words[x].ToString().Contains("/ /") || words[x].ToString().Contains("/  /") || words[x].ToString().Contains("/   /") || words[x].ToString().Contains("/     /")  || words[x].ToString().Equals("/             /")
                                                    || words[x].ToString().Equals("/ ") || words[x].ToString().Equals(" /") || words[x].ToString().Equals("/    /"))
                                                {
                                                    for (int twophoto = 1; twophoto <= 2; twophoto++)
                                                    {
                                                        photoname = lower.ToString() + "_" + i.ToString() + ".jpg";

                                                        if (twophoto == 2)
                                                            words[x] += photoname.ToString();
                                                        else
                                                            words[x] = photoname.ToString();
                                                      //  MessageBox.Show(words[x].ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
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
                                                        //    MessageBox.Show(words[x].ToString(), "Test1", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                                                        savephoto(photoname, count, doc, app);
                                                        i++;
                                                        count++;
                                                    }


                                                }
                                                else if (words[x].ToString().Contains("圖示/") || words[x].ToString().Contains("OK /") || words[x].ToString().Equals("/") || words[x].ToString().Equals(" /") || words[x].ToString().Equals("  /") || words[x].ToString().Equals("   /") || words[x].ToString().Equals("    /")
                                                               ||words[x].ToString().Equals("/ "))
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


                                                    Console.WriteLine(photoname.ToString(), "Test", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
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


                        if (rowPos >= 3+testadd && !(Item.ToString().Contains("作業編號")))
                            {
                                

                                if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                             
                                string sql = @"INSERT INTO `" + lower.ToString() + @"`(`TaskCardName`,`Items`,`TaskCardCode`,`TaskDetail`,`StandardValue`,`Note`,`TaskImage`,`TaskImagepath`) VALUES
                               (@test1,@test2,@test3,@test4,@test5,@test6,@Image,@Imagepath) ";

                                    using (MySqlCommand cmd = new MySqlCommand(sql, conn))//匯入檔案到資料庫
                                    {

                                        cmd.Parameters.Add("@test1", nametask.ToString());
                                        cmd.Parameters.Add("@test2", Item.ToString());
                                        cmd.Parameters.Add("@test3", lower.ToLower().ToString());

                                        if (rowPos >= 3)
                                            cmd.Parameters.Add("@test4", Detail1.ToString());
                                        else
                                            cmd.Parameters.Add("@test4", Detail.ToString());
                                        cmd.Parameters.Add("@test5", standard.ToString());
                                        cmd.Parameters.Add("@test6", Note.ToString());
                                        cmd.Parameters.Add("@Image", lower.ToString() + "_" + Item.ToString());
                                        cmd.Parameters.Add("@Imagepath", "ftp://163.18.57.239/" + lower.ToString() + "_" + Item.ToString());
                                    index = cmd.ExecuteNonQuery();
                                    cmd.Dispose();
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
            try { Microsoft.Office.Interop.Word.Table nowTable = doc.Tables[4]; 
                  
                        for (int rowPos = 3; rowPos <= nowTable.Rows.Count; rowPos++)//抓第幾----
                    {  int title = 0;
                                int change9_6 = 0;
                                int outland = 0;
                            string[] allinput = new string[9];
                       for (int columPos = 1; columPos <= nowTable.Columns.Count; columPos++)//抓底幾個|||

                            {
if (pgbShow.Value >= 450) pgbShow.Value = 450;
else    pgbShow.Value++;
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
                                                        if (Good.ToString().Equals("檢查項目")|| Good.ToString().Equals("檢 查 標 準"))
                                                        {
                                                            title++;
                                                        }
                                                        tt = Good.ToString();//崩潰
                                                        taskdetail = Good.ToString();

                                                        break;
                                                    case 4:
                                                        if (Good.ToString().Equals("檢查標準")|| Good.ToString().Equals("檢 查 標 準"))
                                                        {
                                                            title++;
                                                        }
                                                        st = Good.ToString();//崩潰
                                                        stand = Good.ToString();
                                                        if (st.ToString().Equals("□"))
                                                        {
                                                            st = "";stand = "";
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
                                                        if (Good.ToString().Equals("檢查項目")|| Good.ToString().Equals("檢 查 項 目"))
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
                                                        { ch = ""; checkresult = "";
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
                                                            checkresult = checkresult +"\t"+ stand + "\t"  + maybeone;
                                                            stand = "";
                                                        }
                                                        else if(changindex ==1)//表格蘭有9時
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
                                                        if (!(maybeone.ToString().Equals("□")||stand.ToString().Contains("僅紀錄")))
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
                            if (title==0)
                        {
                            if (!(equipment.ToString().Equals("")))
                            {    if (conn.State != ConnectionState.Open) conn.Open();//連線器打開
                                string sql = @"INSERT INTO `" + lower.ToString().ToLower() + "-f1" + @"`(`TaskCardCode`,`TaskCardName`,`Items`,`TaskDetail`,`Equipment`,`CheckResult`,`StandardValue`,`Note`) VALUES
                                                       (@test1,@test2,@test3,@test4,@test5,@test6,@test7,@test8) ";
                                using (MySqlCommand cmd = new MySqlCommand(sql, conn))//匯入檔案到資料庫
                                {
                                            if (checkresult.ToString().Equals("□"))
                                                checkresult = "";
                                    cmd.Parameters.Add("@test1", lower.ToString().ToLower());
                                    cmd.Parameters.Add("@test2", nametask.ToString());
                                    cmd.Parameters.Add("@test3", Item.ToString());//Item
                                    cmd.Parameters.Add("@test5", equipment.ToString());
                                    cmd.Parameters.Add("@test4", taskdetail.ToString());
                                    
                                    cmd.Parameters.Add("@test6", checkresult.ToString());

                                    cmd.Parameters.Add("@test7", stand.ToString());
                                    cmd.Parameters.Add("@test8", note.ToString());

                                    index = cmd.ExecuteNonQuery();


                                }
                                    Item = ""; equipment = ""; taskdetail = ""; checkresult = ""; stand = ""; note = "";
                                    
                                }
                            
                        }
                            title = 0;
                       


                    }//每列
                        }
                        catch (Exception ex) {  MessageBox.Show(ex.ToString()+"此表單沒有CHECKLIST"); }
                    }


            }//抓表格 第幾個表格
            if (index > 0)
                MessageBox.Show("匯入成功", "資料庫提醒", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);

            doc.Close();
            textBox1.Text = "";
            }
        }
        ////////////////
        private void FTPUpload(string photoname)
        {

            string username = "test1";
            string password = "1234";
            string uploadUrl = "ftp://163.18.57.239/" + lower.ToString() + "/" + photoname;

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
        ///抓WORD圖片

        private void savephoto(string photoname, int count, Microsoft.Office.Interop.Word._Document doc, Microsoft.Office.Interop.Word._Application app)
        {
            //Microsoft.Office.Interop.Word._Application app = new Microsoft.Office.Interop.Word.Application();
            // Microsoft.Office.Interop.Word._Document doc = null;
            //foreach (Microsoft.Office.Interop.Word.InlineShape ish in doc.InlineShapes)
            // {
            // if ((doc.InlineShapes[count].Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapeLinkedPicture) || (doc.InlineShapes[count].Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture))

            //MessageBox.Show(photoname.ToString());
            try { doc.InlineShapes[count].Select(); }
            catch { count--;
                doc.InlineShapes[count].Select();
            }
                app.Selection.Copy();
                Image image = Clipboard.GetImage();//取得檔案中圖片
                pictureBox1.Image = image;
                string fs = @"D:\SCD\TaskPhoto\" + photoname.ToString();
                pictureBox1.Image.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);
                FTPUpload(photoname.ToString());


            
            // }
        }

        private void TaskCheckWord_Load(object sender, EventArgs e)
        {
            this.StyleManager = metroStyleManager1;
            metroStyleManager1.Style = (MetroFramework.MetroColorStyle)Convert.ToInt32(style);
        }
    }
}
