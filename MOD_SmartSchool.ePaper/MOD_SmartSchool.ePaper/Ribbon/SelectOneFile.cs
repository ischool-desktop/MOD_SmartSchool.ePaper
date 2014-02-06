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
        public SelectOneFile()
        {
            InitializeComponent();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇檔案";
            ofd.Filter = "Word檔案 (*.doc)|*.doc";
            ofd.Multiselect = false; //不可多選檔案
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //開啟檔案
                try
                {
                    Document doc = new Document(ofd.FileName);

                    if (checkBoxX1.Checked)
                    {
                        //系統編號
                        Update_ePaper ue = new Update_ePaper(doc, textBoxX1.Text, PrefixStudent.系統編號);
                        if (ue.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
                        {
                            MsgBox.Show("電子報表已上傳!!");
                        }
                        else
                        {
                            MsgBox.Show("已取消上傳電子報表");
                        }
                    }
                    else
                    {
                        //學號
                        Update_ePaper ue = new Update_ePaper(doc, textBoxX1.Text, PrefixStudent.學號);
                        if (ue.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
                        {
                            MsgBox.Show("電子報表已上傳!!");
                        }
                        else
                        {
                            MsgBox.Show("已取消上傳電子報表");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MsgBox.Show("發生錯誤!!\n" + ex.Message);
                }

                ////身分證號
                //Update_ePaper ue = new Update_ePaper(doc, textBoxX1.Text, PrefixStudent.身分證號);
                //ue.ShowDialog();

            }
            else
            {
                MessageBox.Show("未選擇檔案!!");
            }

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
