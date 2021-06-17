using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
    static class Program
    {
        
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
      
       
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
          //  Application.Run(new TaskCheckWord("teste","01","1"));
            Application.Run(new Form1());
           // Application.Run(new Lock());
            try
            {
                //設定應用程式處理異常方式：ThreadException處理
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                //處理UI執行緒異常
                Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
                //處理非UI執行緒異常
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

                #region 應用程式的主入口點
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
                #endregion
            }
            catch (Exception ex)
            {
                string str = GetExceptionMsg(ex, string.Empty);
                MessageBox.Show(str, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);

             
            }
        }
        static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            string str = GetExceptionMsg(e.Exception, e.ToString());
            MessageBox.Show(str, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
         
           
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            string str = GetExceptionMsg(e.ExceptionObject as Exception, e.ToString());
            MessageBox.Show(str, "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
  
           
        }

        /// <summary>
        /// 生成自定義異常訊息
        /// </summary>
        /// <param name="ex">異常物件</param>
        /// <param name="backStr">備用異常訊息：當ex為null時有效</param>
        /// <returns>異常字串文字</returns>
        static string GetExceptionMsg(Exception ex, string backStr)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("****************************異常文字****************************");
            sb.AppendLine("【出現時間】：" + DateTime.Now.ToString());
            if (ex != null)
            {
                sb.AppendLine("【異常型別】：" + ex.GetType().Name);
                sb.AppendLine("【異常資訊】：" + ex.Message);
             //   sb.AppendLine("【堆疊呼叫】：" + ex.StackTrace);
            }
            else
            {
                sb.AppendLine("【未處理異常】：" + backStr);
            }
            sb.AppendLine("***************************************************************");
            return sb.ToString();
        }

    }
}
