using System;
using Microsoft.Vbe.Interop;
namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodeModule
{
    public interface ICodeModuleWrapper
    {
        void AddFromFile(string FileName);
        void AddFromString(string String);
        CodePane CodePane { get; }
        int CountOfDeclarationLines { get; }
        int CountOfLines { get; }
        int CreateEventProc(string EventName, string ObjectName);
        void DeleteLines(int StartLine, int Count = 1);
        bool Find(string Target, ref int StartLine, ref int StartColumn, ref int EndLine, ref int EndColumn, bool WholeWord = false, bool MatchCase = false, bool PatternSearch = false);
        string get_Lines(int StartLine, int Count);
        int get_ProcBodyLine(string ProcName, Microsoft.Vbe.Interop.vbext_ProcKind ProcKind);
        int get_ProcCountLines(string ProcName, Microsoft.Vbe.Interop.vbext_ProcKind ProcKind);
        string get_ProcOfLine(int Line, out Microsoft.Vbe.Interop.vbext_ProcKind ProcKind);
        int get_ProcStartLine(string ProcName, Microsoft.Vbe.Interop.vbext_ProcKind ProcKind);
        void InsertLines(int Line, string String);
        string Name { get; set; }
        VBComponent Parent { get; }
        void ReplaceLine(int Line, string String);
        VBE VBE { get; }
    }
    
}
