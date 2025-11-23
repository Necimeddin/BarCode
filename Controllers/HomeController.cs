using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Mvc;
using MoneyExe.Models;
using SixLabors.Fonts;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Net.NetworkInformation;
using Path = System.IO.Path;
using Rectangle = System.Drawing.Rectangle;

namespace MoneyExe.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;

        public HomeController(IWebHostEnvironment env)
        {
            _env = env;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult UploadExcel()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult UploadFile(IFormFile file)
        {
            try
            {
                if (file == null)
                    return BadRequest("File not selected");

                string uploadsFolder = Path.Combine(_env.WebRootPath, "uploads");
                if (!Directory.Exists(uploadsFolder))
                    Directory.CreateDirectory(uploadsFolder);

                string filePath = Path.Combine(uploadsFolder, file.FileName);
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
                string outputFilePath = GenerateBarcodes(filePath);
                var bytes = System.IO.File.ReadAllBytes(outputFilePath);

                return File(
                    bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    Path.GetFileName(outputFilePath)
                );
            }
            catch (Exception ex)
            {

                throw;
            }

        }

        public string GenerateBarcodes(string inputPath)
        {
            var extension = Path.GetExtension(inputPath);
            string outputPath = Path.Combine(Path.GetDirectoryName(inputPath), $"Barcode_{DateTime.Now.ToString("f")}{Path.GetExtension(inputPath)}");

            var wb = new XLWorkbook(inputPath);
            var ws = wb.Worksheet(1);

            // SSCC sütununu avtomatik tapırıq (18 rəqəmli)
            int ssccCol = FindSsccColumn(ws);
            if (ssccCol == -1)
                throw new Exception("SSCC sütunu tapılmadı!");

            // Son sütun (Barcode əlavə olunacaq)
            int barcodeCol = ws.LastColumnUsed().ColumnNumber() + 1;
            ws.Cell(1, barcodeCol).Value = "Barcode";

            // Eyni SSCC üçün eyni şəkil saxlanır
            Dictionary<string, byte[]> barcodeCache = new Dictionary<string, byte[]>();

            // Sətirləri iterasiya edirik
            int lastRow = ws.LastRowUsed().RowNumber();

            for (int row = 1; row <= lastRow; row++)
            {
                string ssccValue = ws.Cell(row, ssccCol).GetString().Trim();

                // 18 rəqəmli SSCC çıxarırıq
                string extracted = ExtractSscc(ssccValue);
                if (string.IsNullOrEmpty(extracted))
                    continue;

                // Barcode cache-dən
                if (!barcodeCache.TryGetValue(extracted, out byte[] imageBytes))
                {
                    imageBytes = GenerateBarcodeImageSafe(extracted);
                    barcodeCache[extracted] = imageBytes;
                }
                int lastColumn = ws.Row(row).LastCellUsed().Address.ColumnNumber;

                // Növbəti boş sütunu tapırıq (son sütundan sonra)
                int targetColumn = lastColumn + 1;


                // Şəkli Excel-ə əlavə edirik
                using (var ms = new MemoryStream(imageBytes))
                {
                    var img = ws.AddPicture(ms)
                                .MoveTo(ws.Cell(row, targetColumn));

                    // Şəkilin ölçüsünü piksel ilə təyin edirik
                    img.Width = 80;   // 80px genişlik
                    img.Height = 30;  // 30px hündürlük
                }
            }

            // Faylı qeyd edirik
            wb.SaveAs(outputPath);
            return outputPath;
        }

        // =====================
        //   SSCC Column Finder
        // =====================
        private int FindSsccColumn(IXLWorksheet ws)
        {
            int lastCol = ws.LastColumnUsed().ColumnNumber();
            int lastRow = ws.LastRowUsed().RowNumber();

            // 1) Bütün sütunlarda "SSCC" başlığını axtar (bütün sətirlərdə)
            for (int row = 1; row <= Math.Min(10, lastRow); row++) // İlk 10 sətirdə axtar
            {
                for (int col = 1; col <= lastCol; col++)
                {
                    string cellValue = ws.Cell(row, col).GetString().Trim().ToLower();
                    if (cellValue == "sscc")
                        return col;
                }
            }

            // 2) Avtomatik tapma: yalnız TAM 18 rəqəmli xanalar
            for (int col = 1; col <= lastCol; col++)
            {
                int matchCount = 0;
                int checkedRows = 0;

                for (int row = 1; row <= lastRow && checkedRows < 50; row++) // İlk 50 sətirdə yoxla
                {
                    string val = ws.Cell(row, col).GetString().Trim();

                    // TAM 18 rəqəmli SSCC yoxla
                    if (!string.IsNullOrWhiteSpace(val) &&
                        val.Length == 18 &&
                        val.All(char.IsDigit))
                    {
                        matchCount++;
                    }
                    checkedRows++;
                }

                // Ən az 3 uyğunluq varsa — bu SSCC sütunudur
                if (matchCount >= 3)
                    return col;
            }

            return -1;
        }

        // =====================
        //   Extract SSCC
        // =====================
        private string ExtractSscc(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            var trimmed = value.Trim();

            // Yalnız 18 rəqəm olarsa qəbul et
            if (trimmed.Length == 18 && trimmed.All(char.IsDigit))
                return trimmed;

            return null;
        }

        // =====================
        //   Barcode Generator
        // =====================
        private byte[] GenerateBarcodeImageSafe(string sscc)
        {
            if (string.IsNullOrWhiteSpace(sscc))
                throw new ArgumentException("sscc boş ola bilməz", nameof(sscc));
            //sscc = "9460377725072";
            // Scannerdə dəqiq çıxması üçün yalnız SSCC göndərilir (GS1 yoxdur)
            string barcodeData = sscc;

            try
            {
                // Mütləq: Set A istifadə etməyi məcbur edirik (Set C problemlidir)
                var writer = new ZXing.BarcodeWriterPixelData
                {
                    Format = ZXing.BarcodeFormat.ITF,
                    Options = new ZXing.Common.EncodingOptions
                    {
                        Height = 150,
                        Width = 500,
                        Margin = 2,
                        PureBarcode = true
                    }
                };

                // Pixel generasiyası
                var pixelData = writer.Write(barcodeData);

                // PixelData → Bitmap
                using var barcodeBitmap = new Bitmap(pixelData.Width, pixelData.Height, PixelFormat.Format32bppRgb);
                var bitmapData = barcodeBitmap.LockBits(
                    new Rectangle(0, 0, barcodeBitmap.Width, barcodeBitmap.Height),
                    ImageLockMode.WriteOnly,
                    barcodeBitmap.PixelFormat);

                try
                {
                    System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length);
                }
                finally
                {
                    barcodeBitmap.UnlockBits(bitmapData);
                }

                // Alt yazı üçün yer
                int textHeight = 40;
                int totalWidth = barcodeBitmap.Width;
                int totalHeight = barcodeBitmap.Height + textHeight;

                using var finalBmp = new Bitmap(totalWidth, totalHeight, PixelFormat.Format32bppArgb);
                using var g = Graphics.FromImage(finalBmp);

                // Keyfiyyət parametrləri
                g.Clear(Color.White);
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

                // Barcode-u çəkirik
                g.DrawImage(barcodeBitmap, 0, 0, barcodeBitmap.Width, barcodeBitmap.Height);

                // Alt yazı (SSCC)
                var font = new System.Drawing.Font("Arial", 18f, System.Drawing.FontStyle.Bold);
                string textToDraw = sscc;

                // Font ölçüsünü avtomatik tənzimləyirik
                float fontSize = 18;
                SizeF measured = g.MeasureString(textToDraw, font);
                float availableWidth = totalWidth - 4;

                while (measured.Width > availableWidth && fontSize > 8)
                {
                    fontSize--;
                    font.Dispose();
                    font = new System.Drawing.Font("Arial", fontSize, System.Drawing.FontStyle.Bold);
                    measured = g.MeasureString(textToDraw, font);
                }

                RectangleF textRect = new RectangleF(0, barcodeBitmap.Height, totalWidth, textHeight);
                using (System.Drawing.Brush brush = new SolidBrush(Color.Black))
                {
                    g.DrawString(textToDraw, font, brush, textRect, new StringFormat
                    {
                        Alignment = StringAlignment.Center,
                        LineAlignment = StringAlignment.Center
                    });
                }

                font.Dispose();

                using var ms = new MemoryStream();
                finalBmp.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
            catch (Exception ex)
            {
                Console.WriteLine("GenerateBarcodeImageSafe xətası: " + ex.Message);
                throw;
            }
        }
    }
}