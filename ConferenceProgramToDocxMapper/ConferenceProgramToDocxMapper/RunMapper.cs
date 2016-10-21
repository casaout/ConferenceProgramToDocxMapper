

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
        private const string _templateFile = "program.docx"; // word file to define formats (make sure to re-use the style names or change them in Program.cs)

        // output
        private const string _exportFile = "program-processed"; // no file extension! .docx";


        static void Main(string[] args)
        {
            var filePath = GetFilePath();

            // generate word file
            var program = new Program(Path.Combine(filePath, _exportFile), Path.Combine(filePath, _templateFile));

            try
            {
                // get parsed json
                //var json = JsonHelper.GetProgramFromWebsite(_programJsonUri); // directly from website
                var json = JsonHelper.GetProgramFromFile(Path.Combine(filePath, _programJsonFile)); // from file on computer

                // add conference title
                //program.AddSessionTitle(); // TODO: change style + format

                // add icon legend
                program.AddIconLegend();

                var _previousSessionDay = DateTime.MinValue;
                foreach (var session in json.Sessions)
                {
                    // handle day separator
                    var day = DateTime.Parse(session.Day);

                    if (day == DateTime.MinValue || day > _previousSessionDay)
                    {
                        program.AddDaySeparator(day);
                        _previousSessionDay = day;
                    }

                    // if the session is a break
                    if (! session.Type.Equals("Research Papers"))
                    {
                        program.AddBreak(session.Title, session.Time);
                        continue;
                    }
                    // if the session is a paper (add all papers)
                    else
                    {
                        program.AddSessionTitle(session);

                        // in case there are papers for this session, add them
                        if (session.Items != null)
                        {
                            foreach (var paper in session.Items)
                            {
                                foreach (var item in json.Items)
                                {
                                    if (paper.Equals(item.Key))
                                    {
                                        program.AddPaper(item);
                                    }
                                }
                            }
                        }
                    }                    
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
