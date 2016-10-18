

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace ConferenceProgramToDocxMapper
{
    class RunMapper
    {
        // input
        private const string _programJsonUri = "https://www.conference-publishing.com/listJSON.php?Event=FSE16";
        private const string _programJsonFilePath = @"C:\Users\André\Desktop\Yearbook\Program\Testing\program.json";
        
        // output
        private const string _templatePath = @"C:\Users\André\Desktop\Yearbook\Program\Testing\program.docx";
        private const string _exportFilePath = @"C:\Users\André\Desktop\Yearbook\Program\Testing\program-new"; // no file extension! .docx";


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
            jsonString = File.ReadAllText(_programJsonFilePath);

            // Now parse with JSON.Net
            var json = JsonConvert.DeserializeObject<RootObject> (jsonString);




            // generate word file
            var program = new Program(_exportFilePath, _templatePath); // initialize interop word program


            // examples:
            program.AddDaySeparator("Saturday, May 14");
            program.AddSessionTitle("Round Table on Privacy Policies/Protocols", "Sat, May 14, 1o:10 - 10:30", "Ballroom B", "Moderator: Tom Zimmermann");
            program.AddBreak("Morning Break", "Sat, May 14, 10:30 - 11:00");
            program.AddPaper("Raising MSR Researchers: An Experience Report on Teaching a Graduate Seminar Course in Mining Software Repositories (MSR)", "Ahmed E. Hassan (Queen's University, Canada)");
            program.AddPaper("Interactive Exploration of Developer Interaction Traces using a Hidden Markov Model", "Kostadin Damevski, Hui Chen, David Shepherd, and Lori Pollock (Virginia Commonwealth University, USA; Virginia State University, USA; ABB, Inc, USA; University of Delaware, USA)");

            program.SaveAndCloseProgram();
        }

        
    }
}
