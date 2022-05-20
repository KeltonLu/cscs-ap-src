using System.Collections.Generic;
using System.ServiceProcess;
using System.Text;
using System.Windows.Forms;
using System;

namespace CIMSService
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。


        // 全自動的話要把16到29 開啟，要手動的話就把這段註記起來，把35~41開啟
        static void Main()
        {
            ServiceBase[] ServicesToRun;

            // 在同一處理序中可以執行一個以上的使用者服務。若要在這個處理序中
            // 加入另一項服務，請修改下行程式碼，
            // 以建立第二個服務物件。例如，
            //
            //   ServicesToRun = new ServiceBase[] {new Service1(), new MySecondUserService()};
            //
            //ServicesToRun = new ServiceBase[] { new CIMSBatchService() };

            //ServiceBase.Run(ServicesToRun);

            if (Environment.UserInteractive)
            {
                CIMSBatchService s = new CIMSBatchService();

                s.Start(null);

                Console.WriteLine("服務已啟動，請按下 Enter 鍵關閉服務...");
                // 必須要透過 Console.ReadLine(); 先停止程式執行
                // 因為 Windows Service 大多是利用多 Thread 或 Timer 執行長時間的工作
                // 所以雖然主執行緒停止執行了，但服務中的執行緒已經在運行了!
                Console.ReadLine();

                s.Stop();

                Console.WriteLine("服務已關閉");
            }
            else
            {
                ServicesToRun = new ServiceBase[] { new CIMSBatchService() };
                ServiceBase.Run(ServicesToRun);
            }
        }

        //<summary>
        // 应用程序的主入口点。調試時用

        //</summary>
        //[STAThread]
        //static void Main()
        //{
        //    Application.EnableVisualStyles();
        //    Application.SetCompatibleTextRenderingDefault(false);
        //    Application.Run(new FormTest());
        //}

    }
}