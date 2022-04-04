using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PdfCompressWeb.Models;

namespace PdfCompressWeb.Tests
{
    [TestClass]
    public class UnitTest
    {
        /// <summary>
        /// 壓縮pdf檔案功能測試
        /// </summary>
        /// <param name="inputfile"></param>
        [TestMethod]
        [DataRow("ef495ea8-babe-47df-a68b-941047604e7e.pdf")]
        public void CompressPdfTest_ReturnTrue(string inputfile)
        {
            CompressPdf compress = new CompressPdf(true);
            string compressedFilename = compress.CompressFile(inputfile);
        }

        /// <summary>
        /// 轉換pdf至圖片，測試是否有出現獨立目錄，以及數量是否正確
        /// </summary>
        /// <param name="inputfile"></param>
        /// <param name="fileNumbers">圖片數量</param>
        [TestMethod]
        [DataRow("9f2738d6-b338-485a-b0c2-7f590fb6e047.pdf")]
        public void ConvertPdfToImage_CheckImageTotalNumberInFolder_ReturnTrue(string inputfile)
        {
            CompressPdf compress = new CompressPdf(true);
            string outputDictoryName = compress.ConvertPdfToImage(inputfile);
            DirectoryInfo di = new DirectoryInfo(outputDictoryName);
            var files = di.GetFiles();
            Assert.IsTrue(files.Length >0);
        }

        [TestMethod]
        [DataRow("9f2738d6-b338-485a-b0c2-7f590fb6e047")]
        public void CompressImage(string outputDictoryName)
        {
            CompressPdf compress = new CompressPdf(true);
            compress.CompressImage(outputDictoryName);
        }

        [TestMethod]
        [DataRow("9f2738d6-b338-485a-b0c2-7f590fb6e047")]
        public void CreatePdf(string outputDictoryName)
        {
            CompressPdf compress = new CompressPdf(true);
            compress.CreatePdf(outputDictoryName);
        }
    }
}
