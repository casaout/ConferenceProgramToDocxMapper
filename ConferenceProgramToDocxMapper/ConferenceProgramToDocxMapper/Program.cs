using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;
using System.Globalization;
using System.Threading;
using System.Net;

namespace ConferenceProgramToDocxMapper
{
    public class Program
    {
        #region Program Settings

        // define icons per paper type
        private Dictionary<string, string> _iconDictionary = new Dictionary<string, string> { { "Full Paper", "📖" }, { "Invited Talk Paper", "📄" }, { "Invited Talk Abstract", "⚛" }, { "Short Paper", "🔭" } };
        // private Dictionary<string, string> _iconDictionary = new Dictionary<string, string> { { "Journal Paper First", "📖" }, { "Research Paper", "📄" }, { "Industry Paper", "⚛" }, { "Demo Paper", "🔭" } };

        // define style names per style (as used in word template)
        private Dictionary<string, string> _stylesDictionary = new Dictionary<string, string> { {"day", "S Date"}, { "session_title", "S Title" }, { "session_chair", "S Session Chair" }, { "session_break", "S Break" }, { "paper_title", "P Title" }, { "paper_author", "P Author" } };

        private const bool _showWordWhileFilling = false;
        private const string _cultureFormat = "en-US";

        #endregion

        #region Fields

        private Application _wordApplication;
        private Document _word;
        private bool _templateUsed;
        private string _fileSavePath;

        #endregion

        #region Methods

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

            // true: show created word document when program runs
            _wordApplication.Visible = _showWordWhileFilling;

            // set format of conference program culture
            Thread.CurrentThread.CurrentCulture = new CultureInfo(_cultureFormat);

            Console.WriteLine("> Word interop started (settings: ShowWordWhileFilling={0}, Culture={1}, TemplateUsed={2})", _showWordWhileFilling, _cultureFormat, _templateUsed);
        }

        public void SaveWordFile()
        {
            if (File.Exists(_fileSavePath)) File.Delete(_fileSavePath); // replace file
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

            Console.WriteLine("> Finished filling word and closed processor.");
        }

        public void AddDaySeparator(DateTime day)
        {
            AddParagraph("— " + day.ToString("dddd, MMMM d") + " —", GetStyle("day"));
        }

        public void AddSessionTitle(Session session)
        {
            if (string.IsNullOrEmpty(session.Title))
            {
                Console.WriteLine("> WARNING: no session title");
                return;
            }

            var day = DateTime.Parse(session.Day);
            var sessionDayTimeString = day.ToString("ddd, MMM d") + ", " + session.Time;
            var location = (string.IsNullOrEmpty(session.Location)) ? "location unknown" : session.Location;
            var timeLocString = sessionDayTimeString + ", " + location;

            AddParagraph(session.Title, GetStyle("session_title"));
            AddParagraph(timeLocString, GetStyle("session_title"));
            if (! string.IsNullOrEmpty(session.ChairsString)) AddParagraph(session.ChairsString, GetStyle("session_chair"));
        }

        public void AddBreak(string title, string time)
        {
            AddParagraph(title, GetStyle("session_break"));
            AddParagraph(time, GetStyle("session_break"));
        }

        public void AddPaper(Item paper)
        {
            if (string.IsNullOrEmpty(paper.Title))
            {
                Console.WriteLine("> WARNING: no paper title");
                return;
            }

            var titleString = GetIcon(paper.Type) + " " + paper.Title;
            AddParagraph(titleString, GetStyle("paper_title"));
            var authorString = string.Format("{0} ({1})", paper.PersonsString, paper.AffiliationsString);
            AddParagraph(authorString, GetStyle("paper_author"));
        }

        private string GetIcon(string type)
        {
            if (_iconDictionary.ContainsKey(type))
            {
                return _iconDictionary[type];
            }
            else
            {
                return string.Empty; // no icon
            }
        
            // other idea: add image as icon (https://msdn.microsoft.com/en-us/library/ms178792.aspx)
        }

        private object GetStyle(string key)
        {
            if (_stylesDictionary.ContainsKey(key))
            {
                return _stylesDictionary[key];
            }
            else
            {
                return null; // no style = default style
            }
        }

        public void AddIconLegend()
        {
            var legend = "Legend: ";

            foreach (var icon in _iconDictionary)
            {
                legend += string.Format("{0}: {1}, ", icon.Value, icon.Key);
            }

            legend = legend.Trim().TrimEnd(','); // remove last comma
            AddParagraph(legend);
        }

        private void AddParagraph(string text, object styleName = null)
        {
            var paragraph = _word.Paragraphs.Add();
            paragraph.Range.Text = FormatString(text);

            if (styleName != null)
            {
                paragraph.set_Style(ref styleName);
            }

            paragraph.Range.InsertParagraphAfter();

            Console.WriteLine("> Added '{0}' with style '{1}'.", text, styleName);
        }

        private string FormatString(string text)
        {
            text = WebUtility.HtmlDecode(text);
            return text.Trim();
        }

        #endregion
    }
}
