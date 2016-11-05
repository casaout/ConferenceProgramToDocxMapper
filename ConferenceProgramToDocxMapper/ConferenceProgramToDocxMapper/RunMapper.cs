

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace ConferenceProgramToDocxMapper
{
    class RunMapper
    {
        // local path for input/output
        private const string _customFilePath = @""; // @"C:\Users\info\Desktop\Testing\";

        // input
        private const string _programJsonUri = "https://www.conference-publishing.com/listJSON.php?Event=FSE16"; // if program is loaded from web
        private const string _programJsonFile = "program-downloaded.json"; // if program is loaded from local file
        private const string _templateFile = "program-template.docx"; // word file to define formats (make sure to re-use the style names or change them in Program.cs)

        // output
        private const string _exportFile = "program-processed"; // no file extension! .docx";

        static void Main(string[] args)
        {
            CreateWordProgram();
            //CreateHtmlProgram();
        }

        private static void CreateWordProgram()
        {
            var filePath = GetFilePath();

            // generate word file
            var program = new ProgramWord(Path.Combine(filePath, _exportFile), Path.Combine(filePath, _templateFile));

            try
            {
                // get parsed json
                //var json = JsonHelper.GetProgramFromWebsite(_programJsonUri); // directly from website
                var json = JsonHelper.GetProgramFromFile(Path.Combine(filePath, _programJsonFile)); // from file on computer

                // add conference title
                program.AddConferenceTitle(json.NameFull);

                // add icon paper legend
                program.AddIconLegend();

                var _previousSessionDay = DateTime.MinValue;
                var _previousSessionEndTime = TimeSpan.MinValue;
                foreach (var session in json.Sessions)
                {
                    // (Exception case) f message from chair
                    if (session.Title.Equals("Message from the Chairs")) continue;

                    // handle session day and time stuff
                    var day = DateTime.Parse(session.Day);
                    var timeString = session.Time; // format 08:30 - 09:00
                    var startTimeString = timeString.Substring(0, timeString.IndexOf('-')).Trim();
                    var endTimeString = timeString.Substring(timeString.IndexOf('-') + 1, 6).Trim();
                    var startTime = TimeSpan.Parse(startTimeString);
                    var endTime = TimeSpan.Parse(endTimeString);

                    // add day separator (if new day)
                    if (day == DateTime.MinValue || day > _previousSessionDay)
                    {
                        program.AddDaySeparator(day);
                    }

                    // (Exception case) if the session is a break
                    if (session.Type.Equals("Social"))
                    {
                        program.AddBreak(session.Title, session.Time);
                        program.AddNewLine();
                    }
                    // (Exception case) do not show papers for demo track
                    else if (session.Type.Equals("Tool Demonstrations"))
                    {
                        program.AddBreak("Tool Demonstrations", "", session.Location);
                        // don't add new line, as the next item (break) will do this
                    }
                    // if the session is a paper (add all papers)
                    else
                    {
                        // add break (TODO: still needed?)
                        if (day.Date == _previousSessionDay.Date && // same day
                            _previousSessionEndTime != TimeSpan.MinValue && // not first item
                            _previousSessionEndTime < startTime) // gap between
                        {
                            program.AddBreak(_previousSessionEndTime, startTime);
                            program.AddNewLine();
                        }


                        // (Exception case) show keynotes differently to normal sessions
                        if (session.Type.Equals("Keynote"))
                        {
                            program.AddKeynote(session);
                        }
                        // default case: add session
                        else
                        {
                            program.AddSessionTitle(session);
                        }


                        // in case there are papers for this session, add them
                        if (session.Items != null && session.Items.Count > 0)
                        {
                            foreach (var paper in session.Items)
                            {
                                foreach (var item in json.Items)
                                {
                                    if (paper.Equals(item.Key)) //&& session.Type == item.Track) // for some reason doesn't work for visions-track and keynote-track
                                    {
                                        program.AddPaper(item);
                                        break; // if there are duplicates, only first one
                                    }
                                }
                            }
                        }
                        else
                        {
                            program.AddNewLine();
                        }
                    }

                    _previousSessionEndTime = endTime;
                    _previousSessionDay = day;
                }

                // save word file
                program.SaveWordFile();
            }
            catch (Exception e)
            {
                Console.WriteLine("> ERROR: " + e.Message);
            }
            finally
            {
                if (program != null) program.Close();
            }
        }

        private static void CreateHtmlProgram()
        {
            var filePath = GetFilePath();

            // generate word file
            var program = new ProgramHtml(Path.Combine(filePath, _exportFile), Path.Combine(filePath, _templateFile));

            try
            {
                // get parsed json
                //var json = JsonHelper.GetProgramFromWebsite(_programJsonUri); // directly from website
                var json = JsonHelper.GetProgramFromFile(Path.Combine(filePath, _programJsonFile)); // from file on computer

                // add conference title
                program.AddConferenceTitle(json.NameFull);

                // add icon paper legend
                program.AddIconLegend();

                var _previousSessionDay = DateTime.MinValue;
                var _previousSessionEndTime = TimeSpan.MinValue;
                foreach (var session in json.Sessions)
                {
                    // (Exception case) f message from chair
                    if (session.Title.Equals("Message from the Chairs")) continue;

                    // handle session day and time stuff
                    var day = DateTime.Parse(session.Day);
                    var timeString = session.Time; // format 08:30 - 09:00
                    var startTimeString = timeString.Substring(0, timeString.IndexOf('-')).Trim();
                    var endTimeString = timeString.Substring(timeString.IndexOf('-') + 1, 6).Trim();
                    var startTime = TimeSpan.Parse(startTimeString);
                    var endTime = TimeSpan.Parse(endTimeString);

                    // add day separator (if new day)
                    if (day == DateTime.MinValue || day > _previousSessionDay)
                    {
                        program.AddDaySeparator(day);
                    }

                    // (Exception case) if the session is a break
                    if (session.Type.Equals("Social"))
                    {
                        program.AddBreak(session.Title, session.Time);
                        program.AddNewLine();
                    }
                    // (Exception case) do not show papers for demo track
                    else if (session.Type.Equals("Tool Demonstrations"))
                    {
                        program.AddBreak("Tool Demonstrations", "", session.Location);
                        // don't add new line, as the next item (break) will do this
                    }
                    // if the session is a paper (add all papers)
                    else
                    {
                        // add break (TODO: still needed?)
                        if (day.Date == _previousSessionDay.Date && // same day
                            _previousSessionEndTime != TimeSpan.MinValue && // not first item
                            _previousSessionEndTime < startTime) // gap between
                        {
                            program.AddBreak(_previousSessionEndTime, startTime);
                            program.AddNewLine();
                        }


                        // (Exception case) show keynotes differently to normal sessions
                        if (session.Type.Equals("Keynote"))
                        {
                            program.AddKeynote(session);
                        }
                        // default case: add session
                        else
                        {
                            program.AddSessionTitle(session);
                        }


                        // in case there are papers for this session, add them
                        if (session.Items != null && session.Items.Count > 0)
                        {
                            foreach (var paper in session.Items)
                            {
                                foreach (var item in json.Items)
                                {
                                    if (paper.Equals(item.Key)) //&& session.Type == item.Track) // for some reason doesn't work for visions-track and keynote-track
                                    {
                                        program.AddPaper(item);
                                        break; // if there are duplicates, only first one
                                    }
                                }
                            }
                        }
                        else
                        {
                            program.AddNewLine();
                        }
                    }

                    _previousSessionEndTime = endTime;
                    _previousSessionDay = day;
                }

                // save word file
                program.SaveWordFile();
            }
            catch (Exception e)
            {
                Console.WriteLine("> ERROR: " + e.Message);
            }
            finally
            {
                if (program != null) program.Close();
            }
        }

        /// <summary>
        /// returns the current path to the stored files and output files
        /// uses a custom directory (if given), else the current working space of the project
        /// </summary>
        /// <returns></returns>
        private static string GetFilePath()
        {
            if (string.IsNullOrEmpty(_customFilePath))
            {
                var cwd = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                return Path.Combine(cwd, "Data");
            }
            else
            {
                return _customFilePath;
            }
        }
    }
}
