using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace Word_NET_Core.Pages
{
    public class DocCreator
    {
        public string template;
        public string savePath;
        private Application app;
        private Document doc;

        public DocCreator()
        {
            app = new Application();
            doc = null;
        }

        /// <summary>
        /// Set template ".doc" path and filepath for save ".doc"
        /// </summary>
        /// <param name="templatePath">Path for template ".doc"</param>
        /// <param name="savePath">Path for save file ".doc"</param>
        /// <returns>If template not found or save directory not exist
        /// - return false, else - true</returns>
        public bool SetPaths(string templatePath, string savePath)
        {
            bool fileExist = File.Exists(Path.GetFullPath(templatePath));
            bool dirExist = Directory.Exists(
                Path.GetDirectoryName(savePath)
                );
            if (fileExist && dirExist)
            {
                this.template = Path.GetFullPath(templatePath);
                this.savePath = Path.GetFullPath(savePath);
                return true;
            }
            return false;
        }

        public void ActivateDoc()
        {
            this.doc = this.app.Documents.Open(this.template);
            this.doc.Activate();
        }

        public void Save()
        {
            this.doc.SaveAs2(this.savePath);
        }

        public void DeactivateDoc()
        {
            this.doc.Close();
        }

        public Bookmark GetBookmark(string mark)
        {
            return this.doc.Bookmarks[mark];
        }

    }

    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet()
        {
            DocCreator dc = new DocCreator();

            if (dc.SetPaths(@"tmpd\tmp.docx", @"tmpd\result.docx"))
            {
                try
                {
                    dc.ActivateDoc();

                    Bookmark title = dc.GetBookmark("title");
                    title.Range.Text = "TEST один";

                    Bookmark btable = dc.GetBookmark("btable");

                    btable.Range.Rows.Add();

                    btable.Range.Cells[7].Range.Text = "First TEST Cell";
                    btable.Range.Cells[8].Range.Text = "Second Cell";
                    btable.Range.Cells[9].Range.Text = "Third Cell";

                    btable.Range.Rows.Add();

                    btable.Range.Cells[10].Range.Text = "First3 Cell";
                    btable.Range.Cells[11].Range.Font.Color = WdColor.wdColorRed;
                    btable.Range.Cells[11].Range.Text = "Second3 RED Cell";
                    btable.Range.Cells[12].Range.Text = "Third3 Cell";

                    dc.Save();
                    dc.DeactivateDoc();

                }
                catch (Exception ex)
                {
                    dc.DeactivateDoc();
                }
            }

        }
    }
}
