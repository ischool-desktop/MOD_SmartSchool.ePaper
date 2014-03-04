using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MOD_SmartSchool.ePaper
{
    static class tool
    {
        static public string GetFilter(string format)
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
            else if (format == "pdf")
                filter += "Adobe (*.pdf)|*.pdf";
            filter += "|所有檔案 (*.*)|*.*";
            return filter;
        }

        static public void SaveFile(MemoryStream stream, string filename, string format)
        {
            if (format == "xls")
            {
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(stream);
                wb.Save(filename, Aspose.Cells.SaveFormat.Excel97To2003);
            }
            else if (format == "xlsx")
            {
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(stream);
                wb.Save(filename, Aspose.Cells.SaveFormat.Xlsx);
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
            else if (format == "pdf")
            {
                Aspose.Pdf.Document pdf = new Aspose.Pdf.Document(stream);
                pdf.Save(filename);
            }
        }
    }
}
