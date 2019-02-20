using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using NLog;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Office.Core;

namespace UpgradeDocs
{
    public class BusinessClass
    {
        public async void UpgradeOfficeFiles(List<string> listview,Label cWord,Label cExcel,Label Cppt,TextBlock cComplete,Button btnprocess)
        {
            int w = 0;
            int e = 0;
            int p = 0;
            foreach (string filetoProcess in listview)
            {
                if (filetoProcess.EndsWith("xls") || filetoProcess.EndsWith("Xls"))
                {
                    await Task.Run(() => { ConvertToXLSX(filetoProcess); });
                    
                    e++;
                    cExcel.Content = e;
                }
                else if (filetoProcess.EndsWith("doc") || filetoProcess.EndsWith("doc"))
                {
                    await Task.Run(() => { ConvertToDocx(filetoProcess); });
                    w++;
                    cWord.Content = w;
                }
                else if (filetoProcess.EndsWith("ppt") || filetoProcess.EndsWith("Ppt"))
                {
                    await Task.Run(() => { ConvertToPPTX(filetoProcess); });
                    p++;
                    Cppt.Content = p;
                }
            }

            cComplete.Text = "Process Completed!!!";
            cComplete.Foreground=Brushes.DarkGreen;
            btnprocess.Content = "Process";
            btnprocess.IsEnabled = true;

        }

        public void ConvertToXLSX(string vfileName)
        {
            try
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook wb;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                wb = excel.Workbooks.Open(vfileName, Password: "'", UpdateLinks: false);
                object misValue = System.Reflection.Missing.Value;
                // Excel.Workbook wbook = new Excel.Workbook();

                //DirSearch( lstFiles,vPath);
                //foreach (string filetoProcess in lstFiles.Items)
                //{
                // Workbook workbook = new Workbook();
                try
                    {
                        //if (!IsProtected(vfileName))
                        //{
                       
                        //excel.DisplayAlerts = true;
                    //var xlsxFile = xlsFile + "x";
                        string fileSaveAS = vfileName.Substring(0, vfileName.Length - 3)+"xlsx";
                    //wb.SaveAs(Filename: fileSaveAS , FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook);
                        wb.SaveAs(Filename: fileSaveAS, FileFormat:Excel.XlFileFormat.xlOpenXMLWorkbook);
                        wb.Close();
                        excel.Quit();
                        File.Delete(vfileName);
                        
                    //workbook.LoadFromFile(vfileName);
                    //workbook.SaveToFile(vfileName.Substring(0, vfileName.Length - 3) + ".xlsx", ExcelVersion.Version2016);
                    LogManager.GetLogger(typeof(BusinessClass).FullName).Info("Spreadsheet Converted Successfully: " + vfileName);
                    //}
                    //else
                    //    {
                    //        LogManager.GetLogger(typeof(BusinessClass).FullName).Warn("This spreadsheet is protected: " + vfileName );
                            

                    //}
                    }
                    catch (Exception exception)
                    {
                    //wb.Close(false, misValue, misValue);
                    wb.Close();
                    excel.Quit();
                    LogManager.GetLogger(typeof(BusinessClass).FullName).Error("Error while converting the excel document " + vfileName + " Error " + exception.Message);

                    }


                //
                //excel.Quit();
            }
            catch (Exception exception)
            {
                LogManager.GetLogger(typeof(BusinessClass).FullName).Error("Error is directory search, Error: " + exception.Message);
            }
        }

        public void ConvertToDocx(string vFileName)
        {
            try
            {
                Word.Application word = new Word.Application();
                Word.Document doc = new Word.Document();

                //DirSearch(lstFiles,vPath);
                //foreach (string filetoProcess in lstFiles.Items)
                //{

                    try
                    {
                        object missing = System.Reflection.Missing.Value;
                        object save_changes = false;
                        object read_only = true;
                        object add_to_recent_files = false;
                        object confirm_conversions = false;
                        object format = 0;
                    if (!IsProtected(vFileName))
                    {
                        word.Visible = false;
                        //Debug.Print(x + Environment.NewLine)
                        //Method 1 for converting
                        //doc = word.Documents.Open(vFileName);
                        //doc.Convert();
                        //doc.Close();
                        //Method 1 for converting
                        doc =
                            word.Documents.Open(vFileName,
                                ref confirm_conversions, ref read_only,
                                ref add_to_recent_files, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref format, ref missing,
                                false, ref missing, ref missing,
                                ref missing, ref missing);

                        // Save as a .docx file.
                        string new_filename = vFileName.Substring(0, vFileName.Length - 3) + "docx";
                       
                        object file_format = Word.WdSaveFormat.wdFormatStrictOpenXMLDocument;
                        doc.SaveAs(new_filename, ref file_format,
                            ref missing, ref missing, add_to_recent_files,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing);

                        // Close the document without prompting.
                        doc.Close(ref save_changes, ref missing,ref missing);
                       
                        word.Quit();
                        File.Delete(vFileName);
                        LogManager.GetLogger(typeof(BusinessClass).FullName).Info("File Converted Successfully: " + vFileName);
                        //Code sample for the WordGlue
                        //using (Doc doc = new Doc(filetoProcess))
                        //{
                        //    doc.SaveAs("output.docx");
                        //}
                    }
                    else
                    {
                        LogManager.GetLogger(typeof(BusinessClass).FullName).Warn("Document is protected: " + vFileName);
                        doc.Close();
                        word.Quit();
                    }
                }
                    catch (Exception exception)
                    {
                    LogManager.GetLogger(typeof(BusinessClass).FullName).Error("Error in directory search, Error: " + exception.Message);
                    doc.Close();
                        word.Quit();
                }
                //}
                word.Quit();
            }
            catch (Exception exception)
            {
                LogManager.GetLogger(typeof(BusinessClass).FullName).Error("Error while converting the document " + vFileName + " Error " + exception.Message);
            }
        }
        public void ConvertToPPTX(string vfileName)
        {
            try
            {
                PowerPoint.Application powerpoint = new PowerPoint.Application();
                // Excel.Workbook wbook = new Excel.Workbook();

                //DirSearch( lstFiles,vPath);
                //foreach (string filetoProcess in lstFiles.Items)
                //{
                // Workbook workbook = new Workbook();
                try
                {
                    //if (!IsProtected(vfileName))
                    //{
                        //powerpoint.Visible = MsoTriState.msoFalse;
                        //powerpoint.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                        PowerPoint.Presentation pptx = powerpoint.Presentations.Open(vfileName, MsoTriState.msoTrue
                            , MsoTriState.msoTrue, MsoTriState.msoFalse);
                        pptx.SaveAs(vfileName.Substring(0, vfileName.Length - 3) + "pptx"
                            , PowerPoint.PpSaveAsFileType.ppSaveAsDefault);
                        powerpoint.Quit();
                        File.Delete(vfileName);
                        LogManager.GetLogger(typeof(BusinessClass).FullName).Info("Presentation Converted Successfully: " + vfileName);
                    //}
                    //else
                    //{
                    //    LogManager.GetLogger(typeof(BusinessClass).FullName).Info("This ppt file is protected: " + vfileName);


                    //}
                }
                catch (Exception exception)
                {
                    powerpoint.Quit();
                    LogManager.GetLogger(typeof(BusinessClass).FullName).Error("Error while converting the presentation " + vfileName + " Error " + exception.Message);

                }


                //
                //excel.Quit();
            }
            catch (Exception exception)
            {
                LogManager.GetLogger(typeof(BusinessClass).FullName).Error("Error is directory search, Error: " + exception.Message);
            }
        }

        public void  DirSearch( string path,List<string> clistview)
        {
            // Dim thingies = From file In Directory.GetFiles(path) Where file.EndsWith(".doc") And Not FileAttributes.Hidden Select file
            // Dim thingies = From file In Directory.GetFiles(path) Where Not FileAttributes.Hidden Select file
            
                var thingies = from file in Directory.GetFiles(path)
                    where (file.EndsWith(".doc") | file.EndsWith(".Doc") | file.EndsWith(".xls") |
                           file.EndsWith(".Xls") | file.EndsWith(".Ppt") | file.EndsWith(".ppt")) & !file.Contains("~$")
                    select file;
                foreach (string item in thingies)
                {
                clistview.Add(item);
            }
            
          
            //lstFiles.Items.addAddRange(thingies);
            foreach (string subdir in Directory.GetDirectories(path))
                DirSearch(subdir,clistview);
        }
        public static bool IsProtected(string vfile)
        {
            byte[] bytes = File.ReadAllBytes(vfile);

            string prefix = System.Text.Encoding.Default.GetString(bytes.Take(2).ToArray());

            // Zip and not password protected.
            if (prefix == "PK")
            {
                return false;
            }

            // Office format.
            if (prefix == "ÐÏ")
            {
                // XLS 2003
                if ((bytes.Skip(520).Take(1).ToArray()[0] == 254))
                {
                    return true;
                }

                if ((bytes.Skip(532).Take(1).ToArray()[0] == 47))
                {
                    return true;
                }

                //  DOC 2005
                if ((bytes.Skip(523).Take(1).ToArray()[0] == 19))
                {
                    return true;
                }

                // Guessing
                if (bytes.Length < 2000)
                {
                    return false;
                }

                // DOC/XLS 2007+
                String start = System.Text.Encoding.Default.GetString(bytes.Take(2000).ToArray()).Replace('\0', ' ');

                if (start.Contains("E n c r y p t e d P a c k a g e"))
                {
                    return true;
                }

                return false;
            }

            // Unknown format.
            return false;
        }
    }
}
