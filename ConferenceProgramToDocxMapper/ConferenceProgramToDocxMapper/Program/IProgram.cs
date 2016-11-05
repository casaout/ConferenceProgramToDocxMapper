
using System;

namespace ConferenceProgramToDocxMapper.Program
{
    interface IProgram
    {
        void SaveFile();
        void Close();
        void AddConferenceTitle(string text);
        void AddDaySeparator(DateTime day);
        void AddSessionTitle(Session session, bool isKeynote = false);
        void AddPaper(Item paper);
        void AddNewLine();
    }
}
