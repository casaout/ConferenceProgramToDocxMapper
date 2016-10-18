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
    class RunMapper
    {
        //private const string _path = @"C:\Users\André\Desktop\Yearbook\Program\Testing";
        //private const string _fileName = "program-new"; // .docx
        private const string _templatePath = @"C:\Users\André\Desktop\Yearbook\Program\Testing\program.docx";
        private const string _exportFilePath = @"C:\Users\André\Desktop\Yearbook\Program\Testing\program-new"; // no file extension! .docx";


        static void Main(string[] args)
        {
            var program = new Program(_exportFilePath, _templatePath); // initialize interop word program

            program.AddDaySeparator("Saturday, May 14");
            program.AddSessionTitle("Round Table on Privacy Policies/Protocols", "Sat, May 14, 1o:10 - 10:30", "Ballroom B", "Moderator: Tom Zimmermann");
            program.AddBreak("Morning Break", "Sat, May 14, 10:30 - 11:00");
            program.AddPaper("Raising MSR Researchers: An Experience Report on Teaching a Graduate Seminar Course in Mining Software Repositories (MSR)", "Ahmed E. Hassan (Queen's University, Canada)");
            program.AddPaper("Interactive Exploration of Developer Interaction Traces using a Hidden Markov Model", "Kostadin Damevski, Hui Chen, David Shepherd, and Lori Pollock (Virginia Commonwealth University, USA; Virginia State University, USA; ABB, Inc, USA; University of Delaware, USA)");

            program.SaveAndCloseProgram();
        }

        
    }
}
