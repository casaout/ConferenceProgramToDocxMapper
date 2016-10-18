

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace ConferenceProgramToDocxMapper
{
    class RunMapper
    {
        private const string _filePath = @"C:\Users\André\Desktop\Yearbook\Program\Testing\";

        // input
        private const string _programJsonUri = "https://www.conference-publishing.com/listJSON.php?Event=FSE16";
        private const string _programJsonFile = "program.json";
        
        // output
        private const string _templateFile = "program.docx";
        private const string _exportFile = "program-new"; // no file extension! .docx";


        static void Main(string[] args)
        {
            // generate word file
            var program = new Program(Path.Combine(_filePath, _exportFile), Path.Combine(_filePath, _templateFile));

            try
            {
                // get parsed json
                //var json = JsonHelper.GetProgramFromWebsite(_programJsonUri); // directly from website
                var json = JsonHelper.GetProgramFromFile(Path.Combine(_filePath, _programJsonFile)); // from file on computer

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
                Console.WriteLine("> An error occurred: " + e.Message);
            }
            finally
            {
                if (program != null) program.Close();
            }
        }

        
    }
}
