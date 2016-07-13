using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PdfConverterGUI
{
    public partial class frmMain : Form
    {
        private string strDirectory = string.Empty;
        private PdfDocument document = null;

        private BackgroundWorker worker = new BackgroundWorker();

        public frmMain()
        {
            InitializeComponent();

            strDirectory = Application.StartupPath;

            strDirectory = strDirectory + "\\images\\";
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            if(dlg.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = dlg.FileName;
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (txtFileName.Text.Trim().Length == 0)
            {
                return;
            }

            document = PdfReader.Open(txtFileName.Text);

            CheckDirectory();

            prgMain.Maximum = document.Pages.Count;
            prgMain.Minimum = 0;
            prgMain.Value = 0;

            worker.DoWork += Worker_DoWork;

            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;

            worker.RunWorkerAsync();

            return;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Process.Start(strDirectory);
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            int imageCount = 0;

            Console.WriteLine(document.Pages.Count);

            // Iterate pages
            foreach (PdfPage page in document.Pages)
            {
                // Get resources dictionary
                PdfDictionary resources = page.Elements.GetDictionary("/Resources");
                if (resources != null)
                {
                    // Get external objects dictionary
                    PdfDictionary xObjects = resources.Elements.GetDictionary("/XObject");
                    if (xObjects != null)
                    {
                        ICollection<PdfItem> items = xObjects.Elements.Values;
                        // Iterate references to external objects
                        foreach (PdfItem item in items)
                        {
                            PdfReference reference = item as PdfReference;
                            if (reference != null)
                            {
                                PdfDictionary xObject = reference.Value as PdfDictionary;
                                // Is external object an image?
                                if (xObject != null && xObject.Elements.GetString("/Subtype") == "/Image")
                                {
                                    ExportImage(xObject, ref imageCount);
                                    Console.WriteLine(imageCount);

                                    prgMain.Invoke(new ThreadStart(delegate
                                    {
                                        prgMain.Value = imageCount;
                                    }));
                                }
                            }
                        }
                    }
                }
            }
            //MessageBox.Show(imageCount + " images exported.", "Export Images");
            Console.WriteLine("END");
        }

        void ExportImage(PdfDictionary image, ref int count)
        {
            ExportJpegImage(image, ref count);
        }

        void ExportJpegImage(PdfDictionary image, ref int count)
        {
            // Fortunately JPEG has native support in PDF and exporting an image is just writing the stream to a file.
            byte[] stream = image.Stream.Value;

            FileStream fs = new FileStream(String.Format(strDirectory + "Image{0}.jpeg", count++), FileMode.Create, FileAccess.Write);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(stream);
            bw.Close();
        }

        void CheckDirectory()
        {
            try
            {
                if (Directory.Exists(strDirectory))
                {
                    Directory.Delete(strDirectory, true);
                }

                Directory.CreateDirectory(strDirectory);
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n" + ex.StackTrace;
                Console.WriteLine(err);
            }
        }
    }
}
