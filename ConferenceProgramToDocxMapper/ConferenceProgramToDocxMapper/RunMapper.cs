

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
            var jsonString = string.Empty;

            // download program from conference-publishing.com
            //using (var webClient = new System.Net.WebClient())
            //{
            //    Console.WriteLine("Fetching json from '{0}'.", _programJsonUri);
            //    jsonString = webClient.DownloadString(_programJsonUri);
            //}

            // read from json file
            jsonString = File.ReadAllText(Path.Combine(_filePath, _programJsonFile));

            // Now parse with JSON.Net
            var json = JsonConvert.DeserializeObject<RootObject> (jsonString);

            // generate word file
            var program = new Program(Path.Combine(_filePath, _exportFile), Path.Combine(_filePath, _templateFile));

            foreach (var session in json.Sessions)
            {
                // TODO: handle day separator
                // TODO: handle break

                program.AddSessionTitle(session.Title, session.Day + ", " + session.Time, session.Location, session.ChairsString);

                // in case there are papers for this session, add them
                if (session.Items != null)
                {
                    foreach (var paper in session.Items)
                    {
                        foreach (var item in json.Items)
                        {
                            if (paper.Equals(item.Key))
                            {
                                program.AddPaper(item.Title, item.PersonsString);

                                // TODO: add icon
                            }
                        }
                    }
                }                
            }

            program.SaveAndCloseProgram();
        }

        
    }
}
