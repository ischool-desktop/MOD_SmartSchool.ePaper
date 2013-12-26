using System;
using System.Collections.Generic;
using System.Text;
using FISCA.DSAUtil;

namespace MOD_SmartSchool.ePaper
{
    public class FeatureBase
    {
        public static DSResponse CallService(string serviceName, DSRequest request)
        {
            return FISCA.Authentication.DSAServices.CallService(serviceName, request);          
        }
    }
}
