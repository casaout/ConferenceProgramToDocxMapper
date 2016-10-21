

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
        private const string _filePath = @"C:\Users\info\Desktop\Testing\";

        // input
        private const string _programJsonUri = "https://www.conference-publishing.com/listJSON.php?Event=FSE16"; // if program is loaded from web
        private const string _programJsonFile = "program.json"; // if program is loaded from local file
        private const string _templateFile = "program.docx"; // word file to define formats (make sure to re-use the style names or change them in Program.cs)

        // output
        private const string _exportFile = "program-processed"; // no file extension! .docx";


        static void Main(string[] args)
        {
            // generate word file
            var program = new Program(Path.Combine(_filePath, _exportFile), Path.Combine(_filePath, _templateFile));

            try
            {
                // get parsed json
                //var json = JsonHelper.GetProgramFromWebsite(_programJsonUri); // directly from website
                var json = JsonHelper.GetProgramFromFile(Path.Combine(_filePath, _programJsonFile)); // from file on computer

                // add icon legend
                program.AddIconLegend();

                var _previousSessionDay = DateTime.MinValue;
                foreach (var session in json.Sessions)
                {
                    // TODO: handle break
                    var day = DateTime.Parse(session.Day);

                    if (day == DateTime.MinValue || day > _previousSessionDay)
                    {
                        program.AddDaySeparator(day);
                        _previousSessionDay = day;
                    }

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

        
    }
}
