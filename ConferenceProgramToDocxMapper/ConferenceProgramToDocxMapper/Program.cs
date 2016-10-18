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
    public class Program
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

        public void SaveWordFile()
        {
            _wordApplication.ActiveDocument.SaveAs(_fileSavePath, WdSaveFormat.wdFormatDocument);
            // _wordApplication.ActiveDocument.Save();
        }

        public void Close()
        {
            if (_word != null)
            {
                _word.Close();
            }

            if (_wordApplication != null)
            {
                _wordApplication.Quit();
                Marshal.FinalReleaseComObject(_wordApplication);
            }
        }

        public void AddDaySeparator(DateTime day)
        {
            AddParagraph("— " + day.ToString("dddd, MMMM d") + " —", "S Date");
        }

        public void AddSessionTitle(Session session)
        {
            var day = DateTime.Parse(session.Day);
            var sessionDayTimeString = day.ToString("ddd, MMM d") + ", " + session.Time;
            var location = (string.IsNullOrEmpty(session.Location)) ? "location unknown" : session.Location;
            var timeLocString = sessionDayTimeString + ", " + location;

            AddParagraph(session.Title, "S Title");
            AddParagraph(timeLocString, "S Title");
            if (! string.IsNullOrEmpty(session.ChairsString)) AddParagraph(session.ChairsString, "S Session Chair");
        }

        public void AddBreak(string title, string time)
        {
            AddParagraph(title, "S Break");
            AddParagraph(time, "S Break");
        }

        public void AddPaper(Item paper)
        {
            // TODO: add icon (https://msdn.microsoft.com/en-us/library/ms178792.aspx)

            var titleString = paper.Title.Replace("&quot;", "\"").Trim();
            AddParagraph(titleString, "P Title");
            var authorString = string.Format("{0} ({1})", paper.PersonsString, paper.AffiliationsString);
            AddParagraph(authorString, "P Author");
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
