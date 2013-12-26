using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MOD_SmartSchool.ePaper
{
    class Permissions
    {
        //權限代碼
        public static string 學生電子報表 { get { return "Content0125"; } }

        public static bool 學生電子報表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學生電子報表].Executable;
            }
        }

        public static string 班級電子報表 { get { return "Content0165"; } }

        public static bool 班級電子報表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級電子報表].Executable;
            }
        }

        public static string 教師電子報表 { get { return "Content0195"; } }

        public static bool 教師電子報表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[教師電子報表].Executable;
            }
        }

        public static string 課程電子報表 { get { return "Content0215"; } }

        public static bool 課程電子報表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[課程電子報表].Executable;
            }
        }

        public static string 電子報表管理 { get { return "Content0215.MOD_SmartSchool.ePaper"; } }

        public static bool 電子報表管理權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[電子報表管理].Executable;
            }
        }
    }
}
