//*****************************************
//*  作    者：
//*  功能說明：
//*  創建日期：
//*  修改日期：2021-03-12
//*  修改記錄：新增次月下市預測表匯入 陳永銘
//*****************************************
namespace CIMSService
{
    partial class FormTest
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.timer1 = new System.Timers.Timer();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.btLdap = new System.Windows.Forms.Button();
            this.button13 = new System.Windows.Forms.Button();
            this.button14 = new System.Windows.Forms.Button();
            this.button15 = new System.Windows.Forms.Button();
            // 2021-03-12 新增次月下市預測表匯入 陳永銘
            this.button16 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.timer1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(64, 46);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(85, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "小計檔匯入";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // timer1
            // 
            this.timer1.Interval = 1000D;
            this.timer1.SynchronizingObject = this;
            this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Elapsed);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(64, 75);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(85, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "廠商匯入";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(64, 140);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(85, 23);
            this.button3.TabIndex = 2;
            this.button3.Text = "年度預測檔";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(64, 188);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(85, 23);
            this.button4.TabIndex = 3;
            this.button4.Text = "次月預測檔";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click_1);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(64, 243);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(85, 23);
            this.button5.TabIndex = 4;
            this.button5.Text = "日    結";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click_1);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(64, 302);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(85, 23);
            this.button6.TabIndex = 5;
            this.button6.Text = "廠商物料";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(64, 356);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(85, 23);
            this.button7.TabIndex = 6;
            this.button7.Text = "代制費用異動";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(237, 46);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(85, 23);
            this.button8.TabIndex = 7;
            this.button8.Text = "預測檔拆分";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(237, 93);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(85, 23);
            this.button9.TabIndex = 8;
            this.button9.Text = "每月監控";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click_1);
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(237, 140);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(85, 23);
            this.button10.TabIndex = 9;
            this.button10.Text = "每日監控";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click_1);
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(414, 46);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(101, 23);
            this.button11.TabIndex = 10;
            this.button11.Text = "廠商資料匯入";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(427, 127);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(75, 23);
            this.button12.TabIndex = 11;
            this.button12.Text = "button12";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // btLdap
            // 
            this.btLdap.Location = new System.Drawing.Point(427, 168);
            this.btLdap.Name = "btLdap";
            this.btLdap.Size = new System.Drawing.Size(75, 23);
            this.btLdap.TabIndex = 12;
            this.btLdap.Text = "LDAP";
            this.btLdap.UseVisualStyleBackColor = true;
            this.btLdap.Click += new System.EventHandler(this.btLdap_Click);
            // 
            // button13
            // 
            this.button13.Location = new System.Drawing.Point(64, 111);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(116, 23);
            this.button13.TabIndex = 13;
            this.button13.Text = "廠商匯入替換前";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // button14
            // 
            this.button14.Location = new System.Drawing.Point(278, 188);
            this.button14.Name = "button14";
            this.button14.Size = new System.Drawing.Size(75, 23);
            this.button14.TabIndex = 14;
            this.button14.Text = "button14";
            this.button14.UseVisualStyleBackColor = true;
            // 
            // button15
            // 
            this.button15.Location = new System.Drawing.Point(237, 226);
            this.button15.Name = "button15";
            this.button15.Size = new System.Drawing.Size(100, 23);
            this.button15.TabIndex = 15;
            this.button15.Text = "庫存寄卡單 DM";
            this.button15.UseVisualStyleBackColor = true;
            this.button15.Click += new System.EventHandler(this.button15_Click);
            // 
            // button16
            // 2021-03-12 新增次月下市預測表匯入 陳永銘
            this.button16.Location = new System.Drawing.Point(64, 215);
            this.button16.Name = "button16";
            this.button16.Size = new System.Drawing.Size(121, 23);
            this.button16.TabIndex = 16;
            this.button16.Text = "次月下市預測表匯入";
            this.button16.UseVisualStyleBackColor = true;
            this.button16.Click += new System.EventHandler(this.button16_Click);
            // 
            // FormTest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(773, 470);
            // 2021-03-12 新增次月下市預測表匯入 陳永銘
            this.Controls.Add(this.button16);
            this.Controls.Add(this.button15);
            this.Controls.Add(this.button14);
            this.Controls.Add(this.button13);
            this.Controls.Add(this.btLdap);
            this.Controls.Add(this.button12);
            this.Controls.Add(this.button11);
            this.Controls.Add(this.button10);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "FormTest";
            this.Text = "FormTest";
            this.Load += new System.EventHandler(this.FormTest_Load);
            ((System.ComponentModel.ISupportInitialize)(this.timer1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Timers.Timer timer1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.Button btLdap;
        private System.Windows.Forms.Button button13;
        private System.Windows.Forms.Button button14;
        private System.Windows.Forms.Button button15;
        // 2021-03-12 新增次月下市預測表匯入 陳永銘
        private System.Windows.Forms.Button button16;
    }
}