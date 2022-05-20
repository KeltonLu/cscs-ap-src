//*****************************************
//*  作    者：
//*  功能說明：
//*  創建日期：
//*  修改日期：2021-05-18
//*  修改記錄：修改時間比對 陳永銘
//*****************************************
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Configuration;
using CIMSBatch;
using CIMSBatch.Business;
using CIMSBatch.Public;

namespace CIMSService
{
    public partial class CIMSBatchService : ServiceBase
    {
        public CIMSBatchService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // TODO: 在此加入啟動服務的程式碼。


            TriggerOne.Interval = 1000;
            //TriggerTwo.Interval = 1000;
            TriggerOne.Enabled = true;
            //TriggerTwo.Enabled = true;
            BatchBL Bbl = new BatchBL();
            //服務啟動時獲取批次執行時間

            Bbl.GetTriggerTime();
            LogFactory.Write("Service Start", GlobalString.LogType.OpLogCategory);


        }

        protected override void OnStop()
        {
            // TODO: 在此加入停止服務所需執行的終止程式碼。


            TriggerOne.Enabled = false;
            //TriggerTwo.Enabled = false;
            LogFactory.Write("Service Stopped", GlobalString.LogType.OpLogCategory);
        }


        private void TriggerTwo_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //string NowTime = DateTime.Now.ToString("HH:mm:ss");
            //string TriggerTimeTwo = ConfigurationManager.AppSettings["TriggerTwo"];//取得觸發時間2    
            //InOut001BL io = new InOut001BL();
            //if (NowTime == TriggerTimeTwo && io.CheckBatchStatus(1) == false)
            //{

            //    TriggerOne.Enabled = false;
            //    LogFactory.Write("Start Batch", GlobalString.LogType.OpLogCategory);
            //    io.RunBatch();
            //    LogFactory.Write("End Batch", GlobalString.LogType.OpLogCategory);
            //    TriggerOne.Interval = 1000;
            //    TriggerOne.Enabled = true;

            //}
        }
        /// <summary>
        /// 定時觸發批次程式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TriggerOne_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            DateTime dtNowTime = DateTime.Now;
            string NowTime = dtNowTime.ToString("HH:mm:ss");

            BatchBL bl = new BatchBL();

            if (NowTime == GlobalString.TriggerTime.BatchOne)
            {

                LogFactory.Write("Start Batch One", GlobalString.LogType.OpLogCategory);
                bl.RunBatch(dtNowTime, GlobalString.BatchRID.BatchOne);
                LogFactory.Write("End Batch One", GlobalString.LogType.OpLogCategory);
            }
            else if (NowTime == GlobalString.TriggerTime.BatchTwo)
            {
                LogFactory.Write("Start Batch Two", GlobalString.LogType.OpLogCategory);
                bl.RunBatch(dtNowTime, GlobalString.BatchRID.BatchTwo);
                LogFactory.Write("End Batch Two", GlobalString.LogType.OpLogCategory);

            }
            else if (NowTime == GlobalString.TriggerTime.BatchThree)
            {
                LogFactory.Write("Start Batch Three", GlobalString.LogType.OpLogCategory);
                bl.RunBatch(dtNowTime, GlobalString.BatchRID.BatchThree);
                LogFactory.Write("End Batch Three", GlobalString.LogType.OpLogCategory);

            }
            else if (NowTime == GlobalString.TriggerTime.BatchFour)
            {
                LogFactory.Write("Start Batch Four", GlobalString.LogType.OpLogCategory);
                bl.RunBatch(dtNowTime, GlobalString.BatchRID.BatchFour);
                LogFactory.Write("End Batch Four", GlobalString.LogType.OpLogCategory);

            }
            else if (NowTime == GlobalString.TriggerTime.BatchFive)
            {
                LogFactory.Write("Start Batch Five", GlobalString.LogType.OpLogCategory);
                bl.RunBatch(dtNowTime, GlobalString.BatchRID.BatchFive);
                LogFactory.Write("End Batch Five", GlobalString.LogType.OpLogCategory);

            }
            else if (NowTime == GlobalString.TriggerTime.BatchSix)
            {
                LogFactory.Write("Start Batch Six", GlobalString.LogType.OpLogCategory);
                bl.RunBatch(dtNowTime, GlobalString.BatchRID.BatchSix);
                LogFactory.Write("End Batch Six", GlobalString.LogType.OpLogCategory);

            }
            else if (NowTime == GlobalString.TriggerTime.BatchSeven)
            {
                LogFactory.Write("Start Batch Seven", GlobalString.LogType.OpLogCategory);
                bl.RunBatch(dtNowTime, GlobalString.BatchRID.BatchSeven);
                LogFactory.Write("End Batch Seven", GlobalString.LogType.OpLogCategory);

            }
            //2021-05-18 修改時間比對 陳永銘
            //else if (NowTime == GlobalString.TriggerTime.BatchEight)
            else if (NowTime == ConfigurationManager.AppSettings["BatchEight"])
            {
                LogFactory.Write("Start Batch Eight", GlobalString.LogType.OpLogCategory);
                bl.RunBatch(dtNowTime, GlobalString.BatchRID.BatchEight);
                LogFactory.Write("End Batch Eight", GlobalString.LogType.OpLogCategory);

            }

        }


    }
}
