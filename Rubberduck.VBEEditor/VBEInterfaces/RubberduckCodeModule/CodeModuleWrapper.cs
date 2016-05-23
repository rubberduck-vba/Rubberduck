using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.VBEditor.VBEInterfaces.RubberduckCodeModule
{
    public class CodeModuleWrapper : ICodeModuleWrapper
    {

        private readonly CodeModule _codeModule;
        public CodeModuleWrapper(CodeModule codeModule)
        {
            if (codeModule == null)
            {
                throw new ArgumentNullException("CodeModule cannot be null");
            }
            this._codeModule = codeModule;
        }
        public CodeModule CodeModule { get { return _codeModule; } }

        public void AddFromFile(string FileName)
        {
            _codeModule.AddFromFile(FileName);
        }
        public void AddFromString(string String)
        {
            _codeModule.AddFromString(String);
        }
        public CodePane CodePane { get { return _codeModule.CodePane; } }
        public int CountOfDeclarationLines { get { return _codeModule.CountOfDeclarationLines; } }
        public int CountOfLines { get { return _codeModule.CountOfLines; } }
        public int CreateEventProc(string EventName, string ObjectName)
        {
            return _codeModule.CreateEventProc(EventName, ObjectName);
        }
        public void DeleteLines(int StartLine, int Count = 1)
        {
            _codeModule.DeleteLines(StartLine, Count);
        }
        public bool Find(string Target, ref int StartLine, ref int StartColumn, ref int EndLine, ref int EndColumn, bool WholeWord = false, bool MatchCase = false, bool PatternSearch = false)
        {
            return _codeModule.Find(Target, ref StartLine, ref StartColumn, ref EndLine, ref EndColumn, WholeWord, MatchCase, PatternSearch);
        }
        public string get_Lines(int StartLine, int Count)
        {
            return _codeModule.get_Lines(StartLine, Count);
        }
        public int get_ProcBodyLine(string ProcName, Microsoft.Vbe.Interop.vbext_ProcKind ProcKind)
        {
            return _codeModule.get_ProcBodyLine(ProcName, ProcKind);
        }
        public int get_ProcCountLines(string ProcName, Microsoft.Vbe.Interop.vbext_ProcKind ProcKind)
        {
            return _codeModule.get_ProcCountLines(ProcName, ProcKind);
        }
        public string get_ProcOfLine(int Line, out Microsoft.Vbe.Interop.vbext_ProcKind ProcKind)
        {
            return _codeModule.get_ProcOfLine(Line, out ProcKind);
        }
        public int get_ProcStartLine(string ProcName, Microsoft.Vbe.Interop.vbext_ProcKind ProcKind)
        {
            return _codeModule.get_ProcStartLine(ProcName, ProcKind);
        }
        public void InsertLines(int Line, string String)
        {
            _codeModule.InsertLines(Line, String);
        }
        public string Name { get { return _codeModule.Name; } set { _codeModule.Name = value; } }
        public VBComponent Parent { get { return _codeModule.Parent; } }
        public void ReplaceLine(int Line, string String)
        {
            _codeModule.ReplaceLine(Line, String);
        }
        public VBE VBE { get { return _codeModule.VBE; } }


        #region CodeModule extension methods
        public string GetLines(Selection selection)
        {
            return _codeModule.GetLines(selection);
        }
        public QualifiedSelection? QualifiedSelection
        {
            get { return _codeModule.GetSelection(); }
        }

        public void DeleteLines(Selection selection)
        {
            _codeModule.DeleteLines(selection);
        }
        public void SetSelection(QualifiedSelection selection)
        {
            _codeModule.SetSelection(selection);
        }
        #endregion


    }
}
