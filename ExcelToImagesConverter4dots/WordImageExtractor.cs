using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows;


namespace ExcelToImagesConverter4dots
{
    public class ExcelImageExtractor
    {
        public List<string> ExtractedFilepaths = new List<string>();

        public List<FromToExcelImage> ExtractedFromToExcelImages = new List<FromToExcelImage>();

        public string err = "";

        private object missing = System.Reflection.Missing.Value;
        private object yes = true;
        private object no = false;
        private object oDocuments = null;
        private object doc = null;
        private object Shapes = null;
        private object ShapesCount = null;
        private object Shape = null;


        private object Sections = null;
        private object Headers = null;
        private object HeaderShapes = null;

        public bool ExtractImages(string filepath, string slideranges)
        {
            err = "";

            Image image = null;
            object ExcelAppSelection = null;
            object HeaderRangeShape = null;
            int iHeaderRangeShapesCount = -1;
            object HeaderRangeShapesCount = null;
            object HeaderRangeShapes = null;
            object HeaderRange = null;
            object Header = null;

            try
            {
                string tempfile = System.IO.Path.GetTempFileName() + System.IO.Path.GetExtension(filepath);

                System.IO.File.Copy(filepath, tempfile, true);

                OfficeHelper.CreateExcelApplication();

                oDocuments = OfficeHelper.ExcelApp.GetType().InvokeMember("Documents", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.ExcelApp, null);

                OfficeHelper.ExcelApp.GetType().InvokeMember("ActivePrinter", BindingFlags.SetProperty, null, OfficeHelper.ExcelApp, new object[] { "Microsoft XPS Document Writer" });

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { tempfile });

                System.Threading.Thread.Sleep(200);

                OfficeHelper.ExcelApp.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, OfficeHelper.ExcelApp, new object[] { true} );

                System.Threading.Thread.Sleep(200);

                OfficeHelper.ExcelApp.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, OfficeHelper.ExcelApp, null);

                System.Threading.Thread.Sleep(200);

                doc.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, doc,null);

                System.Threading.Thread.Sleep(200);

                object isMissing = Type.Missing;
                object isPrintToFile=(object)true;
                object oFalse = (object)false;

                string imgfp = tempfile + ".xps";

                object oRange = (object)0;
                object oCopies = (object)1;
                object oPageType = (object)0;
                object oItems = (object)0;
                object oPages = "";

                if (slideranges != string.Empty)
                {
                    oRange = (object)4;
                    oPages = slideranges;
                }


                object[] oParam=new object[] { 
                    
                    isMissing, false, oRange, imgfp, isMissing,isMissing, oItems, oCopies, oPages, oPageType, true, true, 
                            isMissing, false, isMissing, isMissing, isMissing, isMissing};

                doc.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, doc, oParam);

                while (true)
                {
                    object obg=OfficeHelper.ExcelApp.GetType().InvokeMember("BackgroundPrintingStatus", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.ExcelApp, null );                    

                    if ((int)obg==0)
                    {
                        break;
                    }

                    System.Windows.Forms.Application.DoEvents();
                }

                //OfficeHelper.ExcelApp.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, OfficeHelper.ExcelApp, oParam);

                doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doc, null);

                System.Threading.Thread.Sleep(200);

                OfficeHelper.QuitExcelApplication();

                System.Threading.Thread.Sleep(200);

                OfficeHelper.QuitOfficeApplications();

                System.Threading.Thread.Sleep(200);

                if (System.IO.File.Exists(imgfp))
                {
                    XPSToImageConverter.SaveXpsPageToBitmap(imgfp);

                    List<string> lst = XPSToImageConverter.lst;

                    for (int k = 0; k < lst.Count; k++)
                    {
                        string imagefp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory, k + 1);

                        Bitmap bmp = new Bitmap(lst[k]);

                        frmOptions.SaveImage(imagefp, bmp);

                        bmp.Dispose();
                        bmp = null;

                        System.IO.File.Delete(lst[k]);

                        if (frmMain.Instance.FirstOutputDocument == string.Empty)
                        {
                            frmMain.Instance.FirstOutputDocument = imagefp;
                        }
                    }

                    //System.IO.File.Delete(imgfp);
                }
                else
                {
                    throw new Exception("XPS File missing !");
                }
                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Convert Excel File to Images") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }        
    }

    public class FromToExcelImage
    {
        public string ExcelFilepath = "";
        public string ImageFilepath = "";
        public int ShapeNr = -1;

        public FromToExcelImageTypeEnum FromToExcelImageType = FromToExcelImageTypeEnum.DocumentInlineShape;

        public enum FromToExcelImageTypeEnum
        {
            HeaderInlineShape,
            FooterInlineShape,
            DocumentInlineShape,
            DocumentShape
        }
    }
}
