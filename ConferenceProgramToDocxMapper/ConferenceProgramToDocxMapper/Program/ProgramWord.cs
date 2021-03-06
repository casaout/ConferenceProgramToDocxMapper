﻿using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.IO;
using System.Globalization;
using System.Threading;
using System.Net;
using ConferenceProgramToDocxMapper.Program;

namespace ConferenceProgramToDocxMapper
{
    public class ProgramWord : IProgram
    {
        #region Program Settings

        // define icons per paper type
        private Dictionary<string, string> _iconDictionary = new Dictionary<string, string> { { "Paper with Artifact", "📎" }, { "Distinguished Paper", "🏆" } };
        // { "Full Paper", "📖" }, { "Invited Talk Paper", "📄" }, { "Invited Talk Abstract", "⚛" }, { "Short Paper", "🔭" 
        // { "Journal Paper First", "📖" }, { "Research Paper", "📄" }, { "Industry Paper", "⚛" }, { "Demo Paper", "🔭" }

        private List<string> _distinguishedPaperIds = new List<string> { "fse16main-mainid163-p", "fse16main-mainid221-p", "fse16main-mainid132-p", "fse16main-mainid60-p", "fse16main-mainid242-p", "fse16main-mainid192-p", "fse16main-mainid65-p" };
        private List<string> _paperWithArtifactsIds = new List<string> { "fse16main-mainid146-p", "fse16main-mainid43-p", "fse16main-mainid65-p", "fse16main-mainid241-p", "fse16main-mainid221-p", "fse16main-mainid230-p", "fse16main-mainid129-p", "fse16main-mainid263-p", "fse16main-mainid84-p", "fse16main-mainid169-p", "fse16main-mainid73-p", "fse16main-mainid39-p", "fse16main-mainid16-p", "fse16main-mainid223-p", "fse16main-mainid233-p", "fse16main-mainid247-p", "fse16main-mainid198-p" };

        // define style names per style (as used in word template)
        private Dictionary<string, string> _stylesDictionary = new Dictionary<string, string> { { "title", "D Title" }, { "legend", "D Legend" }, { "day", "S Date" }, { "session_title", "S Title" }, { "session_chair", "S Session Chair" }, { "session_break", "S Break" }, { "paper_title", "P Title" }, { "paper_author", "P Author" } };

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

        public ProgramWord(string fileSavePath, string templatePath = null)
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

        public void SaveFile()
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

        public void AddConferenceTitle(string text)
        {
            var titleString = "Program of the " + text;
            AddParagraph(titleString, GetStyle("title"));
        }

        public void AddDaySeparator(DateTime day)
        {
            AddParagraph("— " + day.ToString("dddd, MMMM d") + " —", GetStyle("day"));
        }

        public void AddSessionTitle(Session session, bool isKeynote = false)
        {
            var sessionTitle = (isKeynote) ? "Keynote" : session.Title;

            if (string.IsNullOrEmpty(sessionTitle))
            {
                Console.WriteLine("> WARNING: no session title");
                return;
            }

            var day = DateTime.Parse(session.Day);
            var sessionDayTimeString = day.ToString("ddd, MMM d") + ", " + session.Time;
            var location = (string.IsNullOrEmpty(session.Location)) ? "location unknown" : session.Location;
            var timeLocString = sessionDayTimeString + ", " + location;

            AddParagraph(sessionTitle, GetStyle("session_title"));
            AddParagraph(timeLocString, GetStyle("session_title"));

            if (! string.IsNullOrEmpty(session.ChairsString) && session.Chairs.Count > 0)
            {
                var sessionChairString = (session.ChairsString.Contains(",")) ? "Session Chairs: " : "Session Chair: "; // (Exception case) json is not formatted well enough, else we could use: (session.Chairs.Count > 1)
                sessionChairString += session.ChairsString;
                AddParagraph(sessionChairString, GetStyle("session_chair"));
            }
        }

        public void AddKeynote(Session session)
        {
            AddSessionTitle(session, true);
        }

        public void AddBreak(TimeSpan start, TimeSpan end, string location = null)
        {
            var titleString = (start >= new TimeSpan(11, 30, 0) && start <= new TimeSpan(14, 30, 0) && end >= new TimeSpan(11, 30, 0) && end <= new TimeSpan(14, 30, 0))
                ? "Lunch Break"
                : "Break";
            var timeString = string.Format("{0} - {1}", start.ToString(@"hh\:mm"), end.ToString(@"hh\:mm"));

            AddBreak(titleString, timeString, location);
        }

        public void AddBreak(string titleString, string timeString = null, string location = null)
        {
            titleString += (string.IsNullOrEmpty(location)) ? string.Empty : ", " + location;
            AddParagraph(titleString, GetStyle("session_break"));
            if (! string.IsNullOrEmpty(timeString)) AddParagraph(timeString, GetStyle("session_break"));
        }

        public void AddPaper(Item paper)
        {
            if (string.IsNullOrEmpty(paper.Title))
            {
                Console.WriteLine("> WARNING: no paper title");
                return;
            }

            var titleString = GetPaperTitleString(paper);
            AddParagraph(titleString, GetStyle("paper_title"));
            var authorString = string.Format("{0} ({1})", paper.PersonsString, paper.AffiliationsString);
            AddParagraph(authorString, GetStyle("paper_author"));
        }

        private string GetPaperTitleString(Item paper)
        {
            var icon = GetIconSpecialCase(paper);
            var titleString = (string.IsNullOrEmpty(icon)) ? paper.Title : icon + " " + paper.Title;

            // create typeString
            var typeString = string.Empty;
            if (paper.Key.Contains("fse16var")) typeString = "Visions and Reflections";
            else if (paper.Key.Contains("fse16ind")) typeString = "Industry";
            else if (paper.Key.Contains("fse16demo")) typeString = "Demo";
            //if (paper.Type.Equals("Journal-First Paper")) typeString = "Journal First";
            //if (paper.Key.Contains("demo")) typeString = "Demo";

            //if (type.Equals("Research Paper") || type.Equals("Doctoral Symposium") || type.Equals("Keynote")) type = string.Empty
            //var type = paper.Track.TrimEnd('s').Trim(); // removing 's' from type as it's given in plural form
            //if (type.Equals("Tool Demonstration")) type = "Demo";
            //if (type.Equals("Research Paper") || type.Equals("Doctoral Symposium") || type.Equals("Keynote")) type = string.Empty;
            //type = type.Replace("Paper", "").Trim(); // remove 'paper', it's sufficient to have 'Short' or 'Full'
            //type = type.Replace("Tool Demonstration", "Demo").Trim(); // not needed
            //type = type.Replace("Abstract", "").Trim(); // not needed


            // add type string if there is any
            return (string.IsNullOrEmpty(typeString)) ? titleString : string.Format("{0} ({1})", titleString, typeString);
        }

        private string GetIconSpecialCase(Item paper)
        {
            var iconString = string.Empty;
            if (_distinguishedPaperIds.Contains(paper.Key)) iconString += GetIcon("Distinguished Paper") + " ";
            if (_paperWithArtifactsIds.Contains(paper.Key)) iconString += GetIcon("Paper with Artifact") + " ";
            return iconString.Trim();
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

            if (_iconDictionary.Count == 0) return;
            var legend = "Legend: ";

            foreach (var icon in _iconDictionary)
            {
                legend += string.Format("{0}: {1}, ", icon.Value, icon.Key);
            }

            legend = legend.Trim().TrimEnd(','); // remove last comma
            AddParagraph(legend, GetStyle("legend"));
        }

        public void AddNewLine()
        {
            AddParagraph(string.Empty, "Standard");
        }

        private void AddParagraph(string text, object styleName = null)
        {
            var paragraph = _word.Paragraphs.Add();
            paragraph.Range.Text = FormatString(text);

            if (styleName != null)
            {
                try
                {
                    paragraph.set_Style(ref styleName);
                }
                catch
                {
                    Console.WriteLine("ERROR: style '{0}' does not exist.", styleName);
                }
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
