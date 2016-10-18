using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;

namespace ConferenceProgramToDocxMapper
{
    public class Program : IDisposable
    {
        private Application _wordApplication;
        private Document _word;
        private bool _templateUsed;
        private string _fileSavePath;

        public Program(string fileSavePath, string templatePath = null)
        {
            _fileSavePath = fileSavePath;
            _wordApplication = new Application();

            // open existing word file (to re-use styles)
            if (templatePath != null && File.Exists(templatePath))
            {
                _word = _wordApplication.Documents.Open(templatePath);
                _templateUsed = true;
            }
            // create new word file
            else
            {
                _word = _wordApplication.Documents.Add();
            }

            _wordApplication.Visible = true; // show created word document when program runs
        }

        public void SaveAndCloseProgram()
        {
            _wordApplication.ActiveDocument.SaveAs(_fileSavePath, WdSaveFormat.wdFormatDocument);
            // _wordApplication.ActiveDocument.Save();
            _word.Close();
        }

        public void Dispose()
        {
            if (_wordApplication != null)
            {
                _wordApplication.Quit();
                Marshal.FinalReleaseComObject(_wordApplication);
            }
            //GC.SuppressFinalize(this);
        }

        public void AddDaySeparator(string text)
        {
            AddParagraph("— " + text + " —", "S Date");
        }

        public void AddSessionTitle(string title, string time, string location, string chair = null)
        {
            if (string.IsNullOrEmpty(location)) location = "location unknown";
            var timeLocString = time + ", " + location;

            AddParagraph(title, "S Title");
            AddParagraph(timeLocString, "S Title");
            if (chair != null) AddParagraph(chair, "S Session Chair");
        }

        public void AddBreak(string title, string time)
        {
            AddParagraph(title, "S Break");
            AddParagraph(time, "S Break");
        }

        public void AddPaper(string title, string authors)
        {
            AddParagraph(title, "P Title");
            AddParagraph(authors, "P Author");
        }

        private void AddParagraph(string text, object styleName = null)
        {
            var paragraph = _word.Paragraphs.Add(); // paragraph.Content.Paragraphs.Add();
            paragraph.Range.Text = text;

            if (styleName != null)
            {
                paragraph.set_Style(ref styleName); //_document.Styles[styleName]);
            }

            paragraph.Range.InsertParagraphAfter();
        }

    }
}
