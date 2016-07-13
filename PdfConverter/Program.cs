﻿using System;
using System.Collections.Generic;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Advanced;
using System.IO;

namespace PdfConverter
{
    class Program
    {
        static private string strDirectory = string.Empty;
        static private string pdfFilename = string.Empty;

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                return;
            }

            pdfFilename = args[0];

            PdfDocument document = PdfReader.Open(pdfFilename);

            CheckDirectory();

            int imageCount = 0;

            Console.WriteLine("TotalPages: " + document.Pages.Count);

            DateTime dtStart = DateTime.Now;
            Console.WriteLine("START: " + dtStart.GetDateTimeFormats()[140]);
            
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

                                    if (imageCount % 100 == 0)
                                    {
                                        Console.WriteLine(imageCount);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            DateTime dtEnd = DateTime.Now;

            Console.WriteLine("END: " + dtEnd.GetDateTimeFormats()[140]);

            TimeSpan ts = dtEnd - dtStart;

            Console.WriteLine("TotalMiliseconds: " + ts.TotalMilliseconds);
            Console.ReadKey();
        }

        static void ExportImage(PdfDictionary image, ref int count)
        {
            string filter = image.Elements.GetName("/Filter");

            switch (filter)
            {
                case "/DCTDecode":
                    ExportJpegImage(image, ref count);
                    break;

                case "/FlateDecode":
                    ExportAsPngImage(image, ref count);
                    break;
            }
        }

        static void ExportJpegImage(PdfDictionary image, ref int count)
        {
            // Fortunately JPEG has native support in PDF and exporting an image is just writing the stream to a file.
            byte[] stream = image.Stream.Value;
            
            FileStream fs = new FileStream(String.Format("./" + strDirectory + "/Image{0}.jpeg", count++), FileMode.Create, FileAccess.Write);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(stream);
            bw.Close();
        }

        static void ExportAsPngImage(PdfDictionary image, ref int count)
        {
            int width = image.Elements.GetInteger(PdfImage.Keys.Width);
            int height = image.Elements.GetInteger(PdfImage.Keys.Height);
            int bitsPerComponent = image.Elements.GetInteger(PdfImage.Keys.BitsPerComponent);
            // TODO: You can put the code here that converts vom PDF internal image format to a Windows bitmap
            // and use GDI+ to save it in PNG format.
            // It is the work of a day or two for the most important formats. Take a look at the file
            // PdfSharp.Pdf.Advanced/PdfImage.cs to see how we create the PDF image formats.
            // We don't need that feature at the moment and therefore will not implement it.
            // If you write the code for exporting images I would be pleased to publish it in a future release
            // of PDFsharp.
        }

        static void CheckDirectory()
        {
            try
            {
                strDirectory = "./images/";

                if (Directory.Exists(strDirectory))
                {
                    Directory.Delete(strDirectory, true);
                }

                Directory.CreateDirectory(strDirectory);
            }
            catch(Exception ex)
            {
                string err = ex.Message + "\r\n" + ex.StackTrace;
                Console.WriteLine(err);
            }
        }
    }
}
