using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelToImagesConverter4dots
{
    class OfficeHelper
    {        
        public static object ExcelApp = null;        
        public static Type ExcelApplicationType;               

        public static void CreateExcelApplication()
        {
            if (ExcelApp != null)
            {
                try
                {
                    ExcelApp.GetType().InvokeMember("Visible", BindingFlags.IgnoreReturn | BindingFlags.Public |
                    BindingFlags.Static | BindingFlags.SetProperty, null, ExcelApp, new object[] { false }, System.Globalization.CultureInfo.InvariantCulture);
                }
                catch
                {
                    ExcelApplicationType = System.Type.GetTypeFromProgID("Excel.Application");
                    ExcelApp = Activator.CreateInstance(ExcelApplicationType);
                }
            }
            else
            {
                ExcelApplicationType = System.Type.GetTypeFromProgID("Excel.Application");
                ExcelApp = Activator.CreateInstance(ExcelApplicationType);
            }
        }
        public static void QuitExcelApplication()
        {
            if (ExcelApp != null)
            {
                try
                {
                    ExcelApp.GetType().InvokeMember("Quit", BindingFlags.IgnoreReturn | BindingFlags.Instance |
                    BindingFlags.InvokeMethod, null, ExcelApp, null, System.Globalization.CultureInfo.InvariantCulture);
                }
                catch { }

                ExcelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();                
            }
        }

        public static void QuitOfficeApplications()
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            proc.StartInfo.FileName = "QuitOfficeApplications.exe";
            proc.StartInfo.CreateNoWindow = true;

            proc.Start();
            proc.WaitForExit();
        }


    }
}
