using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Configuration;
using System.IO;
using System.Collections;
using System.Drawing;
using System.Text.RegularExpressions;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocumentBarcodeSplitter
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.config = new AppSettingsReader();
            this.fileSystemWatcher.Path = (String)config.GetValue("sourcePath", typeof(String));
            this.fileSystemWatcher.EnableRaisingEvents = true;
        }

        protected override void OnStop()
        {
        }

        private void fileSystemWatcher_Created(object sender, System.IO.FileSystemEventArgs e)
        {
            try
            {
                this.Process(e.FullPath);
            }
            catch (Exception ex)
            {
                this.eventLog.WriteEntry(ex.Message+"\n"+ex.StackTrace, EventLogEntryType.Error);
            }
        }

        private void Process(String filePath)
        {
            ArrayList documentInfos = GetDocumentInfos(filePath);

            if (documentInfos != null)
            {
                PdfDocument inputDocument = PdfReader.Open(filePath, PdfDocumentOpenMode.Import);

                String destinationPath = (String)config.GetValue("destPath", typeof(String));

                foreach (DocumentInfo info in documentInfos)
                {
                    PdfDocument outputDocument = new PdfDocument();
                    outputDocument.Version = inputDocument.Version;

                    for (int page = info.startPage; page <= info.endPage; page++)
                    {
                        outputDocument.AddPage(inputDocument.Pages[page]);
                    }

                    String destinationName = destinationPath + info.barcode + ".pdf";
                    outputDocument.Save(destinationName);
                }

                File.Delete(filePath);
            }
        }

        private ArrayList GetDocumentInfos(String filePath)
        {
            ArrayList documentInfos = new ArrayList();

            // convert PDF to JPEG using ImageMagick
            String tempPath = (String)config.GetValue("tempPath", typeof(String));

            TimeSpan t = (DateTime.UtcNow - new DateTime(1970, 1, 1));
            int timestamp = (int)t.TotalSeconds;

            String fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            String tempFilePath = tempPath + timestamp + ".bmp";

            String imageMagickPath = (String)config.GetValue("imageMagickPath", typeof(String));

            try
            {
                Process convert = new Process();
                convert.StartInfo.FileName = imageMagickPath + "convert.exe";
                convert.StartInfo.Arguments = "-density 200 " + filePath + " " + tempFilePath;
                convert.Start();
                convert.WaitForExit();
            }
            catch (Exception ex)
            {
                this.eventLog.WriteEntry(ex.Message + "\n" + ex.StackTrace, EventLogEntryType.Error);
                return null;
            }

            int index = 0;
            DocumentInfo documentInfo = null;

            do
            {
                String file = tempPath + timestamp + "-" + index + ".bmp";
                if (!File.Exists(file) && index == 0)
                {
                    file = tempPath + timestamp + ".bmp";
                }

                if (File.Exists(file))
                {
                    String barcode = this.ReadBarcode(file);

                    if (barcode.Length > 0)
                    {
                        documentInfo = new DocumentInfo(index);
                        documentInfo.barcode = barcode;
                        documentInfos.Add(documentInfo);
                    }
                    else if (documentInfo == null)
                    {
                        documentInfo = new DocumentInfo(index);
                        documentInfo.barcode = fileName;
                        documentInfos.Add(documentInfo);
                    }
                    else
                    {
                        documentInfo.endPage = documentInfo.endPage + 1;
                    }

                    File.Delete(file);

                }

                index++;
            }
            while (File.Exists(tempPath + timestamp + "-" + index + ".bmp"));

            if (documentInfos.Count == 0)
            {
                this.eventLog.WriteEntry("No temporary file generated for " + fileName, EventLogEntryType.Information);
            }

            return documentInfos;
        }

        private String ReadBarcode(String filePath)
        {
            // Read File and scan for Barcode
            ArrayList barcodes = new ArrayList();
            Bitmap image = (Bitmap)Bitmap.FromFile(filePath);

            BarcodeImaging.FullScanPage(ref barcodes, image, 100);

            image.Dispose();

            if (barcodes.Count > 0)
            {
                Regex regex = new Regex((String)config.GetValue("barcodeRegExp", typeof(String)));
                String barcode = "";

                foreach (String currentBarcode in barcodes)
                {
                    if (regex.IsMatch(currentBarcode))
                    {
                        barcode = currentBarcode.Substring(0, currentBarcode.Length - 1);
                        break;
                    }
                }

                if (barcode.Length > 0)
                {
                    return barcode;
                }
            }

            return "";
        }

        private AppSettingsReader config;
    }
}
