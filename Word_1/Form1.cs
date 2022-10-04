using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word_1.word;
using Word = Microsoft.Office.Interop.Word;

namespace Word_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
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
