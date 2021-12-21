﻿using System;
using System.Collections.Generic;

using System.Text;

using System.Drawing;
using System.Reflection;
using System.Threading;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows;
using System.Windows.Media.Imaging;

namespace ExcelToImagesConverter4dots
{
    public class ExcelImageExtractor
    {
        public List<string> ExtractedFilepaths = new List<string>();

        public List<FromToWordImage> ExtractedFromToWordImages = new List<FromToWordImage>();

        public string err = "";

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
            object oDocuments = null;
            object doc = null;
            object Sections = null;

            object pImage = null;
            object pImageImage = null;

            try
            {
                OfficeHelper.CreateExcelApplication();

                oDocuments = OfficeHelper.ExcelApp.GetType().InvokeMember("Workbooks", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.ExcelApp, null);

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { filepath });

                /*
                System.Threading.Thread.Sleep(100);
                                 
                OfficeHelper.ExcelApp.GetType().InvokeMember("Activate", BindingFlags.IgnoreReturn | BindingFlags.Public |
                BindingFlags.Static | BindingFlags.InvokeMethod, null, OfficeHelper.ExcelApp, null);
                */
                System.Threading.Thread.Sleep(200);

                Sections = doc.GetType().InvokeMember("Sheets", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, null);

                object SectionsCount = Sections.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Sections, null);
                int iSectionsCount = (int)SectionsCount;

                StringRange sr = new StringRange(slideranges);

                for (int m1 = 1; m1 <= iSectionsCount; m1++)
                {
                    if (!sr.IsInRange(m1)) continue;

                    object oSlide = doc.GetType().InvokeMember("Sheets", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, new object[] { m1 });

                    oSlide.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, oSlide, null);

                    object oUsedRange = null;

                    if (slideranges != string.Empty)
                    {
                        string[] ranges = slideranges.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                        for (int k = 0; k < ranges.Length; k++)
                        {
                            StringRange sr2 = new StringRange(ranges[k]);

                            if (!sr2.IsInRange(m1)) continue;

                            int sq1pos = ranges[k].IndexOf(":");
                            int sq2pos = -1;

                            if (ranges[k].Length > sq1pos + 1)
                            {
                                sq2pos = ranges[k].IndexOf(":", sq1pos + 1);
                            }

                            if (sq1pos >= 0 && sq2pos >= 0)
                            {
                                string sq = ranges[k].Substring(sq1pos + 1);

                                oUsedRange = oSlide.GetType().InvokeMember("Range", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oSlide, new object[] { sq });
                            }
                            else if (sq1pos >= 0 && sq2pos < 0)
                            {
                                string sq = ranges[k];

                                oUsedRange = oSlide.GetType().InvokeMember("Range", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oSlide, new object[] { sq });
                            }
                            else
                            {
                                oUsedRange = oSlide.GetType().InvokeMember("UsedRange", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oSlide, null);
                            }

                            object oUsedRangeR = oUsedRange.GetType().InvokeMember("Rows", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oUsedRange, null);
                            object oUsedRangeC = oUsedRange.GetType().InvokeMember("Columns", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oUsedRange, null);

                            object oUsedRangeRCnt = oUsedRangeR.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oUsedRangeR, null);
                            object oUsedRangeCCnt = oUsedRangeC.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oUsedRangeC, null);

                            int iUsedRangeRCnt = (int)oUsedRangeRCnt;
                            int iUsedRangeCCnt = (int)oUsedRangeCCnt;

                            if (iUsedRangeCCnt == 0 && iUsedRangeRCnt == 0)
                            {
                                continue;
                            }

                            object oScreen = (object)1;
                            object oFormat = (object)2;
                            object[] oParam = new object[] { oScreen, oFormat };

                            oUsedRange.GetType().InvokeMember("CopyPicture", BindingFlags.InvokeMethod, null, oUsedRange, oParam);

                            //string imgfp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory, m1);

                            string imgfp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory, ExtractedFilepaths.Count);

                            FromToWordImage wim = new FromToWordImage();
                            wim.WordFilepath = filepath;
                            wim.ImageFilepath = imgfp;

                            Thread thread = new Thread(new ParameterizedThreadStart(SaveInlineShape));
                            thread.SetApartmentState(ApartmentState.STA);
                            thread.Start(wim);
                            thread.Join();

                            oUsedRange = null;
                            oUsedRangeC = null;
                            oUsedRangeCCnt = null;
                            oUsedRangeR = null;
                            oUsedRangeRCnt = null;

                            //ExtractedFilepaths.Add(imgfp);
                        }
                    }
                    else
                    {
                        oUsedRange = oSlide.GetType().InvokeMember("UsedRange", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oSlide, null);

                        object oUsedRangeR = oUsedRange.GetType().InvokeMember("Rows", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oUsedRange, null);
                        object oUsedRangeC = oUsedRange.GetType().InvokeMember("Columns", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oUsedRange, null);

                        object oUsedRangeRCnt = oUsedRangeR.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oUsedRangeR, null);
                        object oUsedRangeCCnt = oUsedRangeC.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oUsedRangeC, null);

                        int iUsedRangeRCnt = (int)oUsedRangeRCnt;
                        int iUsedRangeCCnt = (int)oUsedRangeCCnt;

                        if (iUsedRangeCCnt == 0 && iUsedRangeRCnt == 0)
                        {
                            continue;
                        }

                        object oScreen = (object)1;
                        object oFormat = (object)2;
                        object[] oParam = new object[] { oScreen, oFormat };

                        oUsedRange.GetType().InvokeMember("CopyPicture", BindingFlags.InvokeMethod, null, oUsedRange, oParam);

                        //string imgfp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory, m1);

                        string imgfp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory, ExtractedFilepaths.Count);

                        FromToWordImage wim = new FromToWordImage();
                        wim.WordFilepath = filepath;
                        wim.ImageFilepath = imgfp;

                        Thread thread = new Thread(new ParameterizedThreadStart(SaveInlineShape));
                        thread.SetApartmentState(ApartmentState.STA);
                        thread.Start(wim);
                        thread.Join();

                        //ExtractedFilepaths.Add(imgfp);
                    }

                    oSlide = null;
                }

                Sections = null;
                SectionsCount = null;

                oDocuments = null;
                doc = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Replace Image for Document") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }

        Bitmap GetBitmap(BitmapSource source)
        {
            Bitmap bmp = new Bitmap(
              source.PixelWidth,
              source.PixelHeight,
              PixelFormat.Format32bppPArgb);
            BitmapData data = bmp.LockBits(
              new Rectangle(System.Drawing.Point.Empty, bmp.Size),
              ImageLockMode.WriteOnly,
              PixelFormat.Format32bppPArgb);
            source.CopyPixels(
              Int32Rect.Empty,
              data.Scan0,
              data.Height * data.Stride,
              data.Stride);
            bmp.UnlockBits(data);
            return bmp;
        }

        protected void SaveInlineShape(object owim)
        {
            try
            {
                if (System.Windows.Clipboard.GetDataObject() != null)
                {
                    System.Windows.IDataObject data = System.Windows.Clipboard.GetDataObject();
                    if (data.GetDataPresent(System.Windows.DataFormats.Bitmap))
                    {
                        System.Windows.Interop.InteropBitmap image = (System.Windows.Interop.InteropBitmap)data.GetData(System.Windows.DataFormats.Bitmap, true);

                        Bitmap bmp = GetBitmap(image);

                        //remove for normal

                        /*
                        if (frmAbout.LDT == string.Empty)
                        {
                            using (Graphics g = Graphics.FromImage(bmp))
                            {
                                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                                Font fontp = new Font(frmMain.Instance.Font.FontFamily, 20, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);

                                SizeF sz = g.MeasureString(Module.ApplicationName + " - 4dots Software - Please Register", fontp);

                                g.DrawString(Module.ApplicationName + " - 4dots Software - Please Register", fontp, Brushes.DarkBlue,
                                    new PointF(bmp.Width - sz.Width, bmp.Height - sz.Height));

                                fontp.Dispose();
                            }
                        }
                        */

                        //string imgfp = System.IO.Path.Combine(Module.CurrentImagesDirectory, Guid.NewGuid().ToString() + ".bmp");

                        FromToWordImage wim = owim as FromToWordImage;

                        string imgfp = wim.ImageFilepath;

                        string bmpfp = imgfp + ".bmp";

                        bmp.Save(bmpfp,ImageFormat.Bmp);

                        Bitmap bmp2 = new Bitmap(bmpfp);

                        ExtractedFilepaths.Add(imgfp);                        

                        frmOptions.SaveImage(imgfp, bmp2);

                        if (System.IO.File.Exists(imgfp))
                        {
                            System.IO.FileInfo fi = new System.IO.FileInfo(wim.WordFilepath);
                            System.IO.FileInfo fi2 = new System.IO.FileInfo(imgfp);

                            if (Properties.Settings.Default.KeepCreationDate)
                            {
                                fi2.CreationTime = fi.CreationTime;
                            }

                            if (Properties.Settings.Default.KeepLastModificationDate)
                            {
                                fi2.LastWriteTime = fi.LastWriteTime;
                            }
                        }

                        bmp.Dispose();
                        bmp = null;
                        bmp2.Dispose();
                        bmp2 = null;

                        try
                        {
                            System.IO.File.Delete(bmpfp);
                        }
                        catch
                        { }

                        if (frmMain.Instance.FirstOutputDocument == string.Empty)
                        {
                            frmMain.Instance.FirstOutputDocument = imgfp;
                        }                   

                        ExtractedFromToWordImages.Add(wim);

                        if (frmMain.Instance.FirstOutputDocument == string.Empty)
                        {
                            frmMain.Instance.FirstOutputDocument = imgfp;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }

    public class FromToWordImage
    {
        public string WordFilepath = "";
        public string ImageFilepath = "";
        public int ShapeNr = -1;

        public FromToWordImageTypeEnum FromToWordImageType = FromToWordImageTypeEnum.DocumentInlineShape;

        public enum FromToWordImageTypeEnum
        {
            HeaderInlineShape,
            FooterInlineShape,
            DocumentInlineShape,
            DocumentShape
        }
    }

}