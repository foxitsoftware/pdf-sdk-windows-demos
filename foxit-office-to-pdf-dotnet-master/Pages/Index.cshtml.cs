using System.IO;
using System.Linq;
using System.Threading.Tasks;
using foxit.addon.conversion;
using foxit.common;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Configuration;
using Convert = foxit.addon.conversion.Convert;
using Path = System.IO.Path;

namespace OfficeToPDF.Pages
{
    public class IndexModel : PageModel
    {
        public IFormFile File { get; set; }

        public async Task<IActionResult> OnPost([FromServices] IConfiguration config)
        {
            // Ensure that the uploaded file type is valid.
            var fileExtension = Path.GetExtension(File.FileName).Remove(0, 1).ToLower();
            if (!EnsureExtensionIsConvertible(fileExtension))
            {
                this.ModelState.AddModelError(
                    nameof(this.File), 
                    $"File extension {fileExtension} is not convertible to PDF.");
                return this.Page();
            }
            
            // Save the uploaded file.
            var filePath = $"{Directory.GetCurrentDirectory()}/doc.{fileExtension}";
            using (var stream = System.IO.File.Create(filePath))
            {
                await this.File.CopyToAsync(stream);
            }

            // Return the converted file to the web browser.
            return ConvertToPDFFileResult(config, fileExtension, filePath);
        }

        private static IActionResult ConvertToPDFFileResult(IConfiguration config, string fileExtension, string filePath)
        {
            var saveToPath = $"{Directory.GetCurrentDirectory()}/converted.pdf";
            Library.Initialize(
                config.GetValue<string>("Foxit:SN"), 
                config.GetValue<string>("Foxit:Key"));
            
            if (fileExtension == "docx")
            {
                using var settings = new Word2PDFSettingData();
                Convert.FromWord(filePath, string.Empty, saveToPath, settings);
            }
            else if (fileExtension == "pptx")
            {
                using var settings = new PowerPoint2PDFSettingData();
                Convert.FromPowerPoint(filePath, string.Empty, saveToPath, settings);
            }
            else if (fileExtension == "xlsx")
            {
                using var settings = new Excel2PDFSettingData();
                Convert.FromExcel(filePath, string.Empty, saveToPath, settings);
            }
            // ******
            // HTML conversion requires the "htmltopdf" library for Foxit Support:
            // https://developers.foxitsoftware.com/kb/article/developer-guide-foxit-pdf-sdk-net/#html-to-pdf-conversion
            // ******
            else if (fileExtension.ToLower() == "html")
            {
                using var settings = new HTML2PDFSettingData();
                settings.page_height = 640;
                settings.page_width = 900;
                settings.page_mode = HTML2PDFSettingData.HTML2PDFPageMode.e_PageModeSinglePage;
                Convert.FromHTML(filePath, string.Empty, string.Empty, settings, saveToPath, 20);
            }

            return new PhysicalFileResult(saveToPath, "application/pdf");
        }

        private static string[] ValidExtensions = new[] { "docx","pptx", "xlsx", "html" };
        private bool EnsureExtensionIsConvertible(string fileExtension) =>
            ValidExtensions.Contains(fileExtension);
    }
}
