//*****************************************
//*  作    者：
//*  功能說明：
//*  創建日期：
//*  修改日期：2021-03-12
//*  修改記錄：新增次月下市預測表匯入 陳永銘
//*  修改日期：2021-05-18
//*  修改記錄：參數修正 陳永銘
//*****************************************
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using CIMSBatch.Business;
using CIMSBatch;
using CIMSBatch.FTP;
using System.Configuration;
using CIMSClass;
namespace CIMSService
{
    public partial class FormTest : Form
    {
        BatchBL bl = new BatchBL();
        public FormTest()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(61);
                MessageBox.Show("Import 6.1 OK!");

            }
            catch
            {
                MessageBox.Show("Import 6.1 Fail!");
            }
        }

        private void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //BatchBL Bbl = new BatchBL();
			//2021-05-18 參數修正 陳永銘
            //Bbl.RunBatch(DateTime.Now, 2);
        }

        private void FormTest_Load(object sender, EventArgs e)
        {
            //BatchBL Bbl = new BatchBL();
            //Bbl.GetTriggerTime();
            //timer1.Enabled = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(10);
                MessageBox.Show("Split OK!");
            }
            catch
            {
                MessageBox.Show("Split Fail!");
            }

        }

        /// <summary>
        /// 廠商異動檔執行！
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(62);
                MessageBox.Show("Import 6.2 OK!");

            }
            catch
            {
                MessageBox.Show("Import 6.2 Fail!");

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(14);
                MessageBox.Show("4.14 Batch Process OK!");

            }
            catch
            {
                MessageBox.Show("4.14 Batch Process Fail!");
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }



        private void button5_Click_1(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(63);
                MessageBox.Show("Process 6.3 OK!");

            }
            catch
            {
                MessageBox.Show("Process 6.3 Fail!");

            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(62);
                MessageBox.Show("Import 6.2 OK!");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Import 6.2 Fail!");

            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(64);
                MessageBox.Show("Import 6.4 OK!");

            }
            catch
            {
                MessageBox.Show("Import 6.4 Fail!");

            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(65);
                MessageBox.Show("Import 6.5 OK!");

            }
            catch
            {
                MessageBox.Show("Import 6.5 Fail!");

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(410);
                MessageBox.Show("Import 4.10 OK!");

            }
            catch
            {
                MessageBox.Show("Import 4.10 Fail!");

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(524);
                MessageBox.Show("Import 5.2.4 OK!");

            }
            catch
            {
                MessageBox.Show("Import 5.2.4 Fail!");

            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            try
            {

                bl.TestBatch(13);
                MessageBox.Show("4.13 Batch Process OK!");


            }
            catch
            {
                MessageBox.Show("4.13 Batch Process Fail!");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(214);
                MessageBox.Show("廠商資料匯入 OK!");

            }
            catch
            {
                MessageBox.Show("廠商資料匯入 Fail!");
            }
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(14);
                MessageBox.Show("4.14 Batch Process OK!");

            }
            catch
            {
                MessageBox.Show("4.14 Batch Process Fail!");
            }

        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                string ftpPath = ConfigurationManager.AppSettings["FTPSpecialProjectFilesPath"] + "/" + "TNP";
                FTPFactory ftp = new FTPFactory("Special");
                string[] fileList;

                ArrayList FileNameList = new ArrayList();
                fileList = ftp.GetFileList(ftpPath);
                //ftp.Download(ftpPath, "perso-20090326-TNP.txt", @"d:\cims\", "aa.txt");

                foreach (string str in fileList)
                {
                    MessageBox.Show(str);
                }
            }
            catch
            {
                MessageBox.Show("抓檔失敗!");
            }
        }

        private void btLdap_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(42);
                MessageBox.Show("42 Batch Process OK!");

            }
            catch
            {
                MessageBox.Show("42 Batch Process Fail!");
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(67);
                MessageBox.Show("Import 6.7 OK!");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Import 6.7 Fail!");

            }
            //CIMSClass.Business.InOut007BL bl007 = new CIMSClass.Business.InOut007BL();


        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(68);
                MessageBox.Show("68 OK!");

            }
            catch (Exception ex)
            {
                MessageBox.Show("68 Fail!");
            }
        }

        // 2021-03-12 新增次月下市預測表匯入 陳永銘
        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                bl.TestBatch(66);
                MessageBox.Show("Import 6.6 OK!");
            }
            catch
            {
                MessageBox.Show("Import 6.6 Fail!");
            }
        }
    }


}