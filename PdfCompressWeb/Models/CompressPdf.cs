using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using Ghostscript.NET;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;

namespace PdfCompressWeb.Models
{
    public class CompressPdf
    {
        /// <summary>
        /// 預設圖片長與寬
        /// </summary>
        const int defaultWidth = 595;
        const int defaultHeight = 842;
        public CompressPdf()
        {
            InitialUploadFolder();
            InitialPageSize();
        }

        private void InitialUploadFolder()
        {
            this.UploadFolder = HttpContext.Current.Server.MapPath("~/UploadFiles/");
        }

        /// <summary>
        /// 初始化時設定要壓縮的圖片產出的大小，預設值為A4大小
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public CompressPdf(int width =defaultWidth, int height =defaultHeight)
        {
            InitialPageSize(width, height);

            InitialUploadFolder();
        }

        private void InitialPageSize(int width = defaultWidth, int height = defaultHeight)
        {
            this.Width = width;
            this.Height = height;
        }

        public CompressPdf(bool ForUnitTest)
        {
            if (ForUnitTest)
            {
                //測試時使用bin資料夾
                this.UploadFolder = "";

                InitialPageSize();
            }
        }

        public string UploadFolder { get; set; }

        public int Width { get; set; }
        public int Height { get; set; }
        /// <summary>
        /// 是否調整圖片大小
        /// </summary>
        public bool ResizePicture { get; set; }
        /// <summary>
        /// 圖片類型
        /// </summary>
        public string ImageType { get; set; }

        /// <summary>
        /// 壓縮傳入的pdf檔，並回傳壓縮後的檔案名稱
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public string CompressFile(string source)
        {
            source = this.UploadFolder + source;
            
            string outputFile = GetNewFileName();
            //檔案名稱與路徑不可帶入中文，api無法解析，將產生錯誤
            string outputDictoryName = ConvertPdfToImage(source);
            //壓縮圖片
            CompressImage(outputDictoryName);
            //產生pdf檔
            string outputPDF = CreatePdf(outputDictoryName);

            return outputPDF;
        }

        /// <summary>
        /// 轉換pdf變成圖片
        /// </summary>
        /// <param name="tempFile"></param>
        /// <returns>匯出圖片的目錄名稱</returns>
        public string ConvertPdfToImage(string tempFile)
        {
            PdfReader pdfReader = new PdfReader(tempFile);
            //取得檔名，並作為後續生成的目錄名稱
            string outImageName = Path.GetFileNameWithoutExtension(tempFile);

            GhostscriptPngDevice dev = new GhostscriptPngDevice(GhostscriptPngDeviceType.Png16);
            dev.GraphicsAlphaBits = GhostscriptImageDeviceAlphaBits.V_4;
            dev.TextAlphaBits = GhostscriptImageDeviceAlphaBits.V_4;
            dev.ResolutionXY = new GhostscriptImageDeviceResolution(290, 290);
            dev.InputFiles.Add(tempFile);
            //產生存放圖片的目錄
            Directory.CreateDirectory(this.UploadFolder+outImageName);

            for (int i = 1; i <= pdfReader.NumberOfPages; i++)
            {
                dev.Pdf.FirstPage = i;
                dev.Pdf.LastPage = i;
                dev.CustomSwitches.Add("-dDOINTERPOLATE");
                string outImageNameSub = i.ToString() + "_.png";
                dev.OutputPath = this.UploadFolder + outImageName + "\\" + outImageNameSub;
                dev.Process();
            }

            return outImageName;
        }

        /// <summary>
        /// 逐一取得圖片並壓縮與縮小圖片
        /// </summary>
        /// <param name="outputDictoryName"></param>
        public void CompressImage(string outputDictoryName)
        {
            DirectoryInfo di = new DirectoryInfo(this.UploadFolder + outputDictoryName);
            var imageFiles = di.GetFiles();

            foreach (var imageFile in imageFiles)
            {
               // Get a bitmap.The using statement ensures objects
               //are automatically disposed from memory after use.
                using (Bitmap bmp1 = new Bitmap(imageFile.FullName))
                {                    
                    ImageCodecInfo jpgEncoder = GetEncoder(GetEncoderByImageType());

                    // Create an Encoder object based on the GUID  
                    // for the Quality parameter category.  
                    System.Drawing.Imaging.Encoder myEncoder =
                        System.Drawing.Imaging.Encoder.Quality;

                    // Create an EncoderParameters object.  
                    // An EncoderParameters object has an array of EncoderParameter  
                    // objects. In this case, there is only one  
                    // EncoderParameter object in the array.  
                    EncoderParameters myEncoderParameters = new EncoderParameters(1);

                    // Save the bitmap as a JPG file with zero quality level compression.  
                    EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 0L);
                    myEncoderParameters.Param[0] = myEncoderParameter;

                    var resizedImg = resizeImage(bmp1, new Size(this.Width, this.Height));
                    resizedImg.Save(imageFile.FullName + "_0."+this.ImageType, jpgEncoder, myEncoderParameters);
                }

                //生成壓縮檔完成後，刪除原始檔，避免原檔一並被加入pdf生成內容中
                imageFile.Delete();
            }
        }

        private ImageFormat GetEncoderByImageType()
        {
            ImageFormat imageFormat = null;
            switch (this.ImageType)
            {
                case "jpeg":
                    imageFormat = ImageFormat.Jpeg;
                    break;
                case "png":
                    imageFormat = ImageFormat.Png;
                    break;                
            }
            return imageFormat;
        }

        /// <summary>
        /// 產生pdf檔
        /// </summary>
        /// <param name="outputDictoryName">產生pdf檔的圖片檔來源目錄，會將目錄內的所有圖片加入</param>
        /// <returns>回傳產生的pdf檔名</returns>
        public string CreatePdf(string outputDictoryName)
        {
            Document document = GetDocument();

            DirectoryInfo di = new DirectoryInfo(this.UploadFolder + outputDictoryName);
            var imageFiles = di.GetFiles();

            string outputPdf = GetNewFileName();
            using (var stream = new FileStream(outputPdf, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                PdfWriter.GetInstance(document, stream);
                document.Open();
                foreach (var imageFile in imageFiles)
                {
                    using (var imageStream = new FileStream(imageFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        var image = iTextSharp.text.Image.GetInstance(imageStream);
                        document.Add(image);
                    }
                }
                document.Close();
            }
            return Path.GetFileName(outputPdf);
        }

        private Document GetDocument()
        {
            if (this.ResizePicture)
                return new Document(new iTextSharp.text.Rectangle(595, 842), 0, 0, 0, 0);
            else
                return new Document();
        }

        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            return codecs.Where(x => x.FormatID == format.Guid).FirstOrDefault();
        }

        /// <summary>
        /// 縮小圖片以降低容量(ResizePicture=True的情況下適用)
        /// </summary>
        /// <param name="imgToResize"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        public System.Drawing.Image resizeImage(System.Drawing.Image imgToResize, Size size)
        {
            if (this.ResizePicture)
                return (System.Drawing.Image)(new Bitmap(imgToResize, size));
            else
                return (System.Drawing.Image)(new Bitmap(imgToResize));
        }

        /// <summary>
        /// 產生一個帶有guid名稱的pdf檔案名稱
        /// </summary>
        /// <returns></returns>
        private string GetNewFileName()
        {            
            return this.UploadFolder + Guid.NewGuid().ToString() + ".pdf";
        }
    }
}