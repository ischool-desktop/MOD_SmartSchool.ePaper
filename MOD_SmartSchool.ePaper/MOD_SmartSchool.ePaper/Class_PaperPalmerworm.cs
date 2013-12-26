using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using FISCA.DSAUtil;
using System.Xml;
using System.IO;
using FISCA.Presentation.Controls;

namespace MOD_SmartSchool.ePaper.Class
{
    [FISCA.Permission.FeatureCode("Content0165", "班級電子報表")]
    public partial class Class_PaperPalmerworm : DetailContentBase
    {
        private BackgroundWorker _loader;
        private string _CurrentID;
        private string _RunningID;
        private string _TargetName; //added by Cloud

        private string ViewerType { get { return "Class"; } }

        bool BkWBool = false;

        public Class_PaperPalmerworm()
        {
            InitializeComponent();
            _loader = new BackgroundWorker();
            _loader.DoWork += new DoWorkEventHandler(_loader_DoWork);
            _loader.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_loader_RunWorkerCompleted);
            SaveButtonVisible = false;
            CancelButtonVisible = false;
            Group = "班級電子報表";
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            Changed();
        }

        private void Changed()
        {
            #region 更新時
            if (this.PrimaryKey != "")
            {
                this.Loading = true;

                if (_loader.IsBusy)
                {
                    BkWBool = true;
                }
                else
                {
                    _loader.RunWorkerAsync(this.PrimaryKey);
                }
            }
            #endregion
        }

        private void _loader_DoWork(object sender, DoWorkEventArgs e)
        {
            string running_id = e.Argument as string;
            K12.Data.ClassRecord classrecord = K12.Data.Class.SelectByID(running_id); //added by Cloud
            _TargetName = classrecord.Name; //added by Cloud
            try
            {
                e.Result = QueryElectronicPaper.GetPaperItemByViewer(ViewerType, running_id).GetContent();
            }
            catch (Exception ex)
            {
                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                e.Result = new DSXmlHelper("BOOM");
            }
        }

        private void _loader_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (BkWBool) //如果有其他的更新事件
            {
                BkWBool = false;
                _loader.RunWorkerAsync(this.PrimaryKey);
                return;
            }

            this.Loading = false;

            if (this.IsDisposed)
                return;

            //this.WaitingPicVisible = false;

            dgEPaper.SuspendLayout();

            //刪除
            dgEPaper.Rows.Clear();

            DSXmlHelper helper = e.Result as DSXmlHelper;
            foreach (XmlElement paper in helper.GetElements("PaperItem"))
            {
                DSXmlHelper paperHelper = new DSXmlHelper(paper);

                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgEPaper,
                    paperHelper.GetText("@ID"),
                    paperHelper.GetText("Format"),
                    paperHelper.GetText("PaperName"),
                    DateTimeFormat(paperHelper.GetText("Timestamp")));
                dgEPaper.Rows.Add(row);
                //row.Tag = paperHelper.GetText("Content");
            }

            dgEPaper.ResumeLayout();
        }

        private string DateTimeFormat(string datetime)
        {
            DateTime tryValue;
            if (DateTime.TryParse(datetime, out tryValue))
                return tryValue.ToString("yyyy/MM/dd HH:mm:ss");
            return "";
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgEPaper.SelectedRows.Count <= 0) return;
            DataGridViewRow row = dgEPaper.SelectedRows[0];

            if (MsgBox.Show("您確定要刪除「" + row.Cells[colPaperName.Index].Value + "」嗎？", "", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            string id = "" + row.Cells[colID.Index].Value;
            try
            {
                String before = row.Cells[colPaperName.Index].Value.ToString(); //added by Cloud
                String date = row.Cells[colTimestamp.Index].Value.ToString(); //added by Cloud
                EditElectronicPaper.DeletePaperItem(id);
                //added by Cloud
                FISCA.LogAgent.ApplicationLog.Log("班級電子報表", "刪除", "班級 " + _TargetName + " 的電子報表「" + before + "」(製表日期:" + date + ")被刪除");
            }
            catch (Exception ex)
            {

                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                MsgBox.Show("刪除電子報表發生錯誤。");
            }

            Changed();
        }

        private void btnDownload_Click(object sender, EventArgs e)
        {
            if (dgEPaper.SelectedRows.Count <= 0) return;
            DataGridViewRow row = dgEPaper.SelectedRows[0];

            string base64string = "";
            try
            {
                DSXmlHelper helper = QueryElectronicPaper.GetPaperItemContentById("" + row.Cells[colID.Index].Value).GetContent();
                base64string = helper.GetText("PaperItem/Content");
            }
            catch (Exception ex)
            {

                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                MsgBox.Show(ex.Message);
                return;
            }

            if (string.IsNullOrEmpty(base64string)) return;

            saveFileDialog1.FileName = "" + row.Cells[colPaperName.Index].Value;
            saveFileDialog1.Filter = GetFilter("" + row.Cells[colFormat.Index].Value);
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                //base64 轉 stream
                MemoryStream stream = new MemoryStream(Convert.FromBase64String(base64string));
                SaveFile(stream, saveFileDialog1.FileName, "" + row.Cells[colFormat.Index].Value);
            }
            catch (Exception ex)
            {

                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                MsgBox.Show("儲存檔案發生錯誤。");
                return;
            }

            System.Diagnostics.Process.Start(saveFileDialog1.FileName);
        }

        private void SaveFile(MemoryStream stream, string filename, string format)
        {
            if (format == "xls")
            {
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
                wb.Open(stream);
                wb.Save(filename, Aspose.Cells.FileFormatType.Excel2003);
            }
            else if (format == "xlsx")
            {
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
                wb.Open(stream);
                wb.Save(filename, Aspose.Cells.FileFormatType.Excel2007Xlsx);
            }
            else if (format == "doc")
            {
                Aspose.Words.Document doc = new Aspose.Words.Document(stream);
                doc.Save(filename, Aspose.Words.SaveFormat.Doc);
            }
            else if (format == "docx")
            {
                Aspose.Words.Document doc = new Aspose.Words.Document(stream);
                doc.Save(filename, Aspose.Words.SaveFormat.Docx);
            }
        }

        private string GetFilter(string format)
        {
            string filter = "";
            if (format == "doc")
                filter += "Word 2003 (*.doc)|*.doc";
            else if (format == "docx")
                filter += "Word 2007 (*.docx)|*.docx";
            else if (format == "xls")
                filter += "Excel 2003 (*.xls)|*.xls";
            else if (format == "xlsx")
                filter += "Excel 2007 (*.xlsx)|*.xlsx";
            filter += "|所有檔案 (*.*)|*.*";
            return filter;
        }
    }
}
