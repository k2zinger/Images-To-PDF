using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace UiPathTeam.ImagesToPDF.Activities
{
    [DisplayName("Images To PDF"), Description("Convert a folder full of images to a PDF document")]
    public class ImagesToPDF : NativeActivity
    {
        #region Properties

        [Category("Input"), Description("File path of the image")]
        [RequiredArgument]
        [OverloadGroup("ImageFile")]
        public virtual InArgument<String> ImageFile{ get; set; }

        [Category("Input"), Description("Folder path of the images")]
        [RequiredArgument]
        [OverloadGroup("ImagesFolder")]
        public virtual InArgument<String> ImagesFolder { get; set; }

        [Category("Input"), Description("File path of the PDF document")]
        [RequiredArgument]
        public InArgument<String> OutputFile { get; set; }

        [Category("Input"), Description("Overwrite the destination PDF document if it exists.  Default: False")]
        public InArgument<Boolean> Overwrite { get; set; } = false;

        private List<string> Files = new List<string>();

        private static readonly string[] FORMATS = new string[] { ".bmp", ".emf", ".exif", ".gif", ".icon", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".wmf" };

        #endregion

        #region NativeActivity

        protected override void Execute(NativeActivityContext context)
        {
            #if NET6_0_OR_GREATER
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            #endif

            ValidateParameters(context);

            ExecuteJob(context);
        }

        #endregion

        #region HelperMethods

        public void ValidateParameters(NativeActivityContext context)
        {
            if (!String.IsNullOrEmpty(ImageFile.Get(context)))
            {
                if (!File.Exists(ImageFile.Get(context)))
                {
                    throw new Exception("ImageFile Not Found: " + ImageFile.Get(context));
                }
                else
                {
                    Console.WriteLine("ImageFile: " + ImageFile.Get(context));
                    Files.Add(ImageFile.Get(context));
                    ImagesFolder.Set(context, Path.GetDirectoryName(ImageFile.Get(context)));
                }
            }
            else
            {
                if (!Directory.Exists(ImagesFolder.Get(context)))
                {
                    throw new Exception("ImagesFolder Not Found: " + ImagesFolder.Get(context));
                }
                else
                {
                    if (Files.Count() == 0)
                    {
                        if ((Files = Directory.EnumerateFiles(ImagesFolder.Get(context)).Where(file => FORMATS.Contains(Path.GetExtension(file).ToLower())).ToList()).Count() == 0)
                        {
                            if (!String.IsNullOrEmpty(ImageFile.Get(context)))
                            {
                                throw new Exception("Not a valid image: " + ImageFile.Get(context));
                            }
                            else
                            {
                                throw new Exception("No images found within ImagesFolder: " + ImagesFolder.Get(context));
                            }
                        }
                        Console.WriteLine(Files.Count().ToString() + " images found within the ImagesFolder: " + ImagesFolder.Get(context));
                    }
                }
            }

            if (!Path.GetExtension(OutputFile.Get(context)).ToLower().Equals(".pdf"))
            {
                throw new Exception("Invalid OutputFile.  Extension must be (.pdf): " + OutputFile.Get(context));
            }

            if (Directory.Exists(Path.GetDirectoryName(OutputFile.Get(context))))
            {
                if (File.Exists(OutputFile.Get(context)))
                {
                    if (Overwrite.Get(context))
                    {
                        try
                        {
                            File.Delete(OutputFile.Get(context));
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Error deleting file: " + OutputFile.Get(context) + " -- " + ex.Message);
                        }
                    }
                    else
                    {
                        throw new Exception("OutputFile exists, but Overwrite = False: " + OutputFile.Get(context));
                    }
                }
                Console.WriteLine("OutputFile: " + OutputFile.Get(context));
            }
            else
            {
                throw new Exception("Path to OutputFile does not exist: " + Path.GetDirectoryName(OutputFile.Get(context)));
            }
        }

        public void ExecuteJob(NativeActivityContext context)
        {
            // Create new pdf document and page
            var doc = new PdfDocument();

            // Dim files As String() = directory.GetFiles(folderpath,"*.png")
            
            foreach (string filename in Files)
            {
                Console.WriteLine("Adding: " + Path.GetFileName(filename));
                // Create new oPage
                PdfPage oPage = new PdfPage();

                // Add the page to the pdf document and add the captured image to it
                doc.Pages.Add(oPage);

                // Dim img As PDFSharp.Drawing.XImage  = PDFSharp.Drawing.XImage.FromFile(filename)

                // Create an XImage Object from Image File 
                XImage xImg = null;
                try
                {
                    xImg = XImage.FromFile(filename);

                    oPage.Width = xImg.PixelWidth * 72 / xImg.HorizontalResolution;
                    oPage.Height = xImg.PixelHeight * 72 / xImg.VerticalResolution;

                    // Draw current image file to page
                    XGraphics xgr = XGraphics.FromPdfPage(oPage);
                    xgr.DrawImage(xImg, 0, 0, oPage.Width, oPage.Height);
                }
                catch (Exception ex)
                {
                    try
                    {
                        xImg.Dispose();
                    }
                    catch (Exception)
                    {
                    }
                    throw new Exception("Error reading image file: " + filename + " -- " + ex.Message);
                }
            }

            try
            {
                Console.WriteLine("Saving PDF Document: " + OutputFile.Get(context));
                doc.Save(OutputFile.Get(context));
            }
            catch (Exception ex)
            {
                throw new Exception("Error saving PDF Document -- " + ex.Message);
            }
            finally
            {
                try
                {
                    doc.Close();
                }
                catch (Exception)
                {
                }
            }
        }

        #endregion

    }
}