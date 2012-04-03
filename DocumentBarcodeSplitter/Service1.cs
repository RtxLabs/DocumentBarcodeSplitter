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
            this.Process(e.FullPath);
        }

        private void Process(String filePath)
        {
            ArrayList documentInfos = GetDocumentInfos(filePath);

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

        private ArrayList GetDocumentInfos(String filePath)
        {
            ArrayList documentInfos = new ArrayList();

            // convert PDF to JPEG using ImageMagick
            String tempPath = (String)config.GetValue("tempPath", typeof(String));

            String fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            String tempFilePath = tempPath + fileName + ".bmp";

            String imageMagickPath = (String)config.GetValue("imageMagickPath", typeof(String));

            Process convert = new Process();
            convert.StartInfo.FileName = imageMagickPath + "\\convert.exe";
            convert.StartInfo.Arguments = "-density 200 " + filePath + " " + tempFilePath;
            convert.Start();
            convert.WaitForExit();

            int index = 0;
            DocumentInfo documentInfo = null;
            while (File.Exists(tempPath + fileName + "-" + index + ".bmp"))
            {
                String barcode = this.ReadBarcode(tempPath + fileName + "-" + index + ".bmp");

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

                File.Delete(tempPath + fileName + "-" + index + ".bmp");
                index++;
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
                Regex regex = new Regex(@"^RE-\d+");
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
