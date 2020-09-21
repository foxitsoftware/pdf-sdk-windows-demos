using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using foxit.addon;
using foxit.common;
using foxit.common.fxcrt;
using foxit.pdf;
using foxit.pdf.annots;
using foxit.pdf.interform;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace PdfRedaction.Pages
{
    public class IndexModel : PageModel
    {
        
        public class UserDataJsonModel
        {
            public string FullName { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public int SIN { get; set; }
        }
        
        /** Found under "lib/gsdk_sn.txt" **/
        private static string serialNo = "GOES_HERE";

        /** Found under "lib/gsdk_key.txt" **/
        private static string key = "GOES_HERE";

        public UserDataJsonModel UserData { get; set; }
        
        public async Task OnGetAsync()
        {
            using var fileStream = System.IO.File.OpenRead("./datasource.json");
            this.UserData = await JsonSerializer.DeserializeAsync<UserDataJsonModel>(fileStream);
        }

        public async Task<IActionResult> OnPostPDFAsync()
        {
            using var fileStream = System.IO.File.OpenRead("./datasource.json");
            UserDataJsonModel model = await JsonSerializer.DeserializeAsync<UserDataJsonModel>(fileStream);
            string path = GeneratePDF(model, false);
            var stream = new FileStream(path, FileMode.Open);
            return new FileStreamResult(stream, "application/pdf");
        }

        public async Task<IActionResult> OnPostSecurePDFAsync()
        {
            using var fileStream = System.IO.File.OpenRead("./datasource.json");
            UserDataJsonModel model = await JsonSerializer.DeserializeAsync<UserDataJsonModel>(fileStream);
            string path = GeneratePDF(model, true);
            var stream = new FileStream(path, FileMode.Open);
            return new FileStreamResult(stream, "application/pdf");
        }

        private static string GeneratePDF(UserDataJsonModel model, bool applyRedaction)
        {
            Library.Initialize(serialNo, key);

            try
            {
                // Create a new PDF with a blank page.
                using PDFDoc doc = new PDFDoc();
                using Form form = new Form(doc);
                using PDFPage page = doc.InsertPage(0, PDFPage.Size.e_SizeLetter);

                // Create a text field for each property of the user data json model.
                int iteration = 0;
                foreach (var prop in model.GetType().GetProperties())
                {
                    int offset = iteration * 60;
                    using Control control = form.AddControl(page, prop.Name, Field.Type.e_TypeTextField,
                        new RectF(50f, 600f - offset, 400f, 640f - offset));
                    using Field field = control.GetField();
                    var propValue = prop.GetValue(model);
                    field.SetValue($"{prop.Name}: {propValue}");
                    iteration++;
                }

                // Convert fillable form fields into readonly text labels.
                page.Flatten(true, (int) PDFPage.FlattenOptions.e_FlattenAll);

                if (applyRedaction)
                {
                    // Configure our text search to look at the first (and only) page in our PDF.
                    using var redaction = new Redaction(doc);
                    page.StartParse((int) foxit.pdf.PDFPage.ParseFlags.e_ParsePageNormal, null, false);
                    using TextPage textPage = new TextPage(page,
                        (int) foxit.pdf.TextPage.TextParseFlags.e_ParseTextUseStreamOrder);
                    using TextSearch search = new TextSearch(textPage);
                    RectFArray rectArray = new RectFArray();

                    // Mark text starting with the pattern "SIN:" to be redacted.
                    search.SetPattern("SIN:");
                    while (search.FindNext())
                    {
                        using var searchSentence = new TextSearch(textPage);
                        var sentence = search.GetMatchSentence();
                        searchSentence.SetPattern(sentence.Substring(sentence.IndexOf("SIN:")));

                        while (searchSentence.FindNext())
                        {
                            RectFArray itemArray = searchSentence.GetMatchRects();
                            rectArray.InsertAt(rectArray.GetSize(), itemArray);
                        }
                    }

                    // If we had matches to redact, then apply the redaction.
                    if (rectArray.GetSize() > 0)
                    {
                        using Redact redact = redaction.MarkRedactAnnot(page, rectArray);
                        redact.SetFillColor(0xFF0000);
                        redact.ResetAppearanceStream();
                        redaction.Apply();
                    }
                }

                // Save the file locally and return the file's location to the caller.
                string fileOutput = "./Output.pdf";
                doc.SaveAs(fileOutput, (int) PDFDoc.SaveFlags.e_SaveFlagNoOriginal);
                return fileOutput;
            }
            finally
            {
                Library.Release();
            }
        }
    }
}