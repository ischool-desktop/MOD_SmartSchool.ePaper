using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using Aspose.Words;
using Campus.ePaper;

namespace MOD_SmartSchool.ePaper
{
    public partial class SelectOneFile : BaseForm
    {
        BackgroundWorker BGW { get; set; }

        public SelectOneFile()
        {
            InitializeComponent();

            BGW = new BackgroundWorker();
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.ProgressChanged += new ProgressChangedEventHandler(BGW_ProgressChanged);
            BGW.WorkerReportsProgress = true;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBoxX1.Text))
            {
                MsgBox.Show("請設定報表名稱!!");
                return;
            }

            btnUpdate.Enabled = false;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇檔案";
            ofd.Filter = "Word檔案 (*.doc)|*.doc";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                this.Text = "學生電子檔案上傳(檔案開啟中...)";
                BGW.RunWorkerAsync(ofd);
            }
            else
            {
                btnUpdate.Enabled = true;
                MsgBox.Show("未選擇檔案!!");
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            OpenFileDialog ofd = (OpenFileDialog)e.Argument;

            List<Document> docList = new List<Document>();

            if (ofd.FileNames.Count() == 1)
            {
                Document doc = new Document(ofd.FileName);
                docList.Add(doc);
            }
            else if (ofd.FileNames.Count() > 1)
            {
                int successCount = 0; //用於計算進度。

                foreach (string each in ofd.FileNames)
                {
                    successCount++;

                    Document doc = new Document(each);
                    docList.Add(doc);

                    decimal seed = (decimal)successCount / ofd.FileNames.Count();
                    BGW.ReportProgress((int)(seed * 100), each);
                }
            }
            else
            {
                return;
            }

            e.Result = docList;
        }

        void BGW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string[] s = e.UserState.ToString().Split('\\');
            if (s.Length > 0)
                FISCA.Presentation.MotherForm.SetStatusBarMessage("開啟檔案:" + s[s.Length - 1], e.ProgressPercentage);
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                if (e.Error == null)
                {

                    List<Document> docList = (List<Document>)e.Result;
                    if (checkBoxX1.Checked)
                    {
                        //系統編號
                        Update_ePaper ue = new Update_ePaper(docList, textBoxX1.Text, PrefixStudent.系統編號);
                        if (ue.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
                        {
                            this.Text = "學生電子檔案上傳(電子報表已上傳!!)";
                            MsgBox.Show("電子報表已上傳!!");
                        }
                        else
                        {
                            this.Text = "學生電子檔案上傳(已取消!!)";
                            MsgBox.Show("已取消!!");
                        }
                    }
                    else
                    {
                        //學號
                        Update_ePaper ue = new Update_ePaper(docList, textBoxX1.Text, PrefixStudent.學號);
                        if (ue.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
                        {
                            this.Text = "學生電子檔案上傳(電子報表已上傳!!)";
                            MsgBox.Show("電子報表已上傳!!");
                        }
                        else
                        {
                            this.Text = "學生電子檔案上傳(已取消!!)";
                            MsgBox.Show("已取消!!");
                        }
                    }
                }
                else
                {
                    this.Text = "學生電子檔案上傳(發生錯誤!!)";
                    MsgBox.Show("發生錯誤!!\n" + e.Error.Message);
                }
            }
            else
            {
                this.Text = "學生電子檔案上傳(作業已中止!!)";
                MsgBox.Show("作業已中止!!");
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
