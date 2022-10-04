using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word_1.word
{
    class DocCreator
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
            if(fileExist && dirExist)
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
}
