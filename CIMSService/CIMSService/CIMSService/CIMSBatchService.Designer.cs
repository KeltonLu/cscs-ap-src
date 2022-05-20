namespace CIMSService
{
    partial class CIMSBatchService
    {
        /// <summary> 
        /// 設計工具所需的變數。

        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。

        /// </summary>
        /// <param name="disposing">如果應該公開 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。

        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.TriggerOne = new System.Timers.Timer();
            this.TriggerTwo = new System.Timers.Timer();
            ((System.ComponentModel.ISupportInitialize)(this.TriggerOne)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TriggerTwo)).BeginInit();
            // 
            // TriggerOne
            // 
            this.TriggerOne.Enabled = true;
            this.TriggerOne.Elapsed += new System.Timers.ElapsedEventHandler(this.TriggerOne_Elapsed);
            // 
            // TriggerTwo
            // 
            this.TriggerTwo.Enabled = true;
            this.TriggerTwo.Elapsed += new System.Timers.ElapsedEventHandler(this.TriggerTwo_Elapsed);
            // 
            // CIMSBatchService
            // 
            this.ServiceName = "CIMSService";
            ((System.ComponentModel.ISupportInitialize)(this.TriggerOne)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TriggerTwo)).EndInit();

        }

        #endregion

        private System.Timers.Timer TriggerOne;
        private System.Timers.Timer TriggerTwo;


        public void Start(string[] args)
        {
            this.OnStart(args);
        }

        public void Stop()
        {
            this.OnStop();
        }
    }
}
