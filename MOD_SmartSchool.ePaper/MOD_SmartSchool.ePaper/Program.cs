﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Presentation;
using FISCA.Permission;

namespace MOD_SmartSchool.ePaper
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void Main()
        {
            //電子報表的提供者。
            //SmartSchool.ePaper.DispatcherProvider.Register("ischool", new DispatcherImp(), true);

            //Aspose.Words.License lic_word = new Aspose.Words.License();
            //System.IO.MemoryStream ms_word = new System.IO.MemoryStream(Properties.Resources.Aspose_Total_201402);
            //lic_word.SetLicense(ms_word);

            //Aspose.Pdf.License lic_pdf = new Aspose.Pdf.License();
            //System.IO.MemoryStream ms_pdf = new System.IO.MemoryStream(Properties.Resources.Aspose_Total_201402);
            //lic_pdf.SetLicense(ms_pdf);

            //Aspose.Pdf.License lic_excel = new Aspose.Pdf.License();
            //System.IO.MemoryStream ms_excel = new System.IO.MemoryStream(Properties.Resources.Aspose_Total_201402);
            //lic_excel.SetLicense(ms_excel);

            //管理功能
            MotherForm.StartMenu["電子報表管理"].Enable = Permissions.電子報表管理權限;
            MotherForm.StartMenu["電子報表管理"].Image = Properties.Resources.mail_ok_64;
            MotherForm.StartMenu["電子報表管理"].Click += delegate
            {
                new ElectronicPaperManager().ShowDialog();
            };

            //New
            RibbonBarItem StudentItem = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "其它"];
            StudentItem["電子報表上傳"].Image = Properties.Resources.files_up_64;
            StudentItem["電子報表上傳"].Enable = Permissions.學生電子報表上傳權限;
            StudentItem["電子報表上傳"].Click += delegate
            {
                SelectOneFile f = new SelectOneFile();
                f.ShowDialog();
            };


            //
            RibbonBarItem TeacherItem = FISCA.Presentation.MotherForm.RibbonBarItems["教師", "其它"];
            TeacherItem["電子報表上傳"].Image = Properties.Resources.files_up_64;
            TeacherItem["電子報表上傳"].Enable = Permissions.教師電子報表上傳權限;
            TeacherItem["電子報表上傳"].Click += delegate
            {
                SelectOneFile_Teacher f = new SelectOneFile_Teacher();
                f.ShowDialog();
            };

            //資料項目

            FeatureAce UserPermission;
            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.學生電子報表];
            if (UserPermission.Editable || UserPermission.Viewable)
                K12.Presentation.NLDPanels.Student.AddDetailBulider(new FISCA.Presentation.DetailBulider<Student.Student_PaperPalmerworm>());

            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.班級電子報表];
            if (UserPermission.Editable || UserPermission.Viewable)
                K12.Presentation.NLDPanels.Class.AddDetailBulider(new FISCA.Presentation.DetailBulider<Class.Class_PaperPalmerworm>());

            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.教師電子報表];
            if (UserPermission.Editable || UserPermission.Viewable)
                K12.Presentation.NLDPanels.Teacher.AddDetailBulider(new FISCA.Presentation.DetailBulider<Teacher.Teacher_PaperPalmerworm>());

            UserPermission = FISCA.Permission.UserAcl.Current[Permissions.課程電子報表];
            if (UserPermission.Editable || UserPermission.Viewable)
                K12.Presentation.NLDPanels.Course.AddDetailBulider(new FISCA.Presentation.DetailBulider<Course.Course_PaperPalmerworm>());


            //New
            Catalog ribbon = RoleAclSource.Instance["學生"]["功能按鈕"];
            ribbon.Add(new RibbonFeature(Permissions.學生電子報表上傳, "電子報表上傳"));

            Catalog detail = RoleAclSource.Instance["學生"]["資料項目"];
            detail.Add(new DetailItemFeature(Permissions.學生電子報表, "學生電子報表"));

            detail = RoleAclSource.Instance["班級"]["資料項目"];
            detail.Add(new DetailItemFeature(Permissions.班級電子報表, "班級電子報表"));

            ribbon = RoleAclSource.Instance["教師"]["功能按鈕"];
            ribbon.Add(new RibbonFeature(Permissions.教師電子報表上傳, "電子報表上傳"));

            detail = RoleAclSource.Instance["教師"]["資料項目"];
            detail.Add(new DetailItemFeature(Permissions.教師電子報表, "教師電子報表"));

            detail = RoleAclSource.Instance["課程"]["資料項目"];
            detail.Add(new DetailItemFeature(Permissions.課程電子報表, "課程電子報表"));

            detail = RoleAclSource.Instance["系統"]["功能按鈕"];
            detail.Add(new ReportFeature(Permissions.電子報表管理, "電子報表管理"));
        }
    }
}
