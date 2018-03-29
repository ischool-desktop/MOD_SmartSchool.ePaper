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
    public partial class SelectOneFile_Teacher : BaseForm
    {
        BackgroundWorker BGW { get; set; }

        //2018/3/28 穎驊註解，本教師電子報表上傳功能，為因應高雄小組會議  [05-03][06] 電子報表發給老師 製作
        public SelectOneFile_Teacher()
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
            if (string.IsNullOrEmpty(textBoxX1.Text.Trim()))
            {
                MsgBox.Show("請設定報表名稱!!");
                return;
            }

            btnUpdate.Enabled = false;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇檔案";
            ofd.Filter = "Word檔案 (*.docx;*.doc)|*.docx;*.doc";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                this.Text = "教師電子檔案上傳(檔案開啟中...)";
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
            btnUpdate.Enabled = true;
            this.Text = "教師電子檔案上傳";
            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    
                    List<Document> docList = (List<Document>)e.Result;
                    if (checkBoxX1.Checked)
                    {
                        //系統編號
                        Update_ePaper_Teacher ue = new Update_ePaper_Teacher(docList, textBoxX1.Text, PrefixTeacher.系統編號);
                        if (ue.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
                        {
                            MsgBox.Show("電子報表已上傳!!");
                        }
                        else
                        {
                            MsgBox.Show("已取消!!");
                        }
                    }
                    //身分字號
                    else if (checkBoxX3.Checked)
                    {
                        Update_ePaper_Teacher ue = new Update_ePaper_Teacher(docList, textBoxX1.Text, PrefixTeacher.身分證號);
                        if (ue.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
                        {
                            MsgBox.Show("電子報表已上傳!!");
                        }
                        else
                        {
                            MsgBox.Show("已取消!!");
                        }
                    }
                    else
                    {
                        // 2018/3/28 穎驊註解，之所以要用這麼麻煩的的Key 值是因為教師並沒有像是學生有著學號可以通用的Key 值可以使用
                        //教師姓名_暱稱{王小明_自然科專任老師}
                        Update_ePaper_Teacher ue = new Update_ePaper_Teacher(docList, textBoxX1.Text, PrefixTeacher.教師姓名_暱稱);
                        if (ue.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
                        {
                            MsgBox.Show("電子報表已上傳!!");
                        }
                        else
                        {
                            MsgBox.Show("已取消!!");
                        }
                    }
                }
                else
                {
                    this.Text = "教師電子檔案上傳(發生錯誤!!)";
                    MsgBox.Show("發生錯誤!!\n" + e.Error.Message);
                }
            }
            else
            {
                this.Text = "教師電子檔案上傳(作業已中止!!)";
                MsgBox.Show("作業已中止!!");
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
