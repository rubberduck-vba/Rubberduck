using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class CodeModule : SafeComWrapper<Microsoft.Vbe.Interop.CodeModule>
    {
        public CodeModule(Microsoft.Vbe.Interop.CodeModule comObject) 
            : base(comObject)
        {
        }

        public void AddFromString(string value)
        {
            Invoke(content => ComObject.AddFromString(content), value);
        }

        public void AddFromFile(string path)
        {
            Invoke(content => ComObject.AddFromFile(content), path);
        }

        public void InsertLines(int line, string value)
        {
            Invoke((i, content) => ComObject.InsertLines(i, content), line, value);
        }

        public void DeleteLines(int startLine, int count = 1)
        {
            Invoke((i, c) => ComObject.DeleteLines(i, c), startLine, count);
        }

        public void ReplaceLine(int line, string value)
        {
            Invoke((i, c) => ComObject.ReplaceLine(i, c), line, value);
        }

        public int CreateEventProc(string eventName, string objectName)
        {
            Invoke((e, o) => ComObject.CreateEventProc(e, o), eventName, objectName);
        }

        public Selection? Find(string target, bool wholeWord = false, bool matchCase = false, bool patternSearch = false)
        {
            return InvokeResult((t, word, caseSensitive, usePattern) =>
            {
                var startLine = 0;
                var startColumn = 0;
                var endLine = 0;
                var endColumn = 0;
                if (ComObject.Find(t, ref startLine, ref startColumn, ref endLine, ref endColumn, word, caseSensitive, usePattern))
                {
                    return new Selection(startLine, startColumn, endLine, endColumn);
                }
                return (Selection?)null;
            }, target, wholeWord, matchCase, patternSearch);
        }

        public VBComponent Parent { get; private set; }
        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        public string Name
        {
            get { return InvokeResult(() => ComObject.Name); }
            set { Invoke(v => ComObject.Name = v, value); }
        }

        public string GetLines(int startLine, int count)
        {
            return InvokeResult((s, c) => ComObject.get_Lines(s, c), startLine, count);
        }

        public int CountOfLines { get { return InvokeResult(() => ComObject.CountOfLines); } }
        public int GetProcStartLine(string procName, ProcKind procKind)
        {
            return InvokeResult((n, k) => ComObject.get_ProcStartLine(n, k), procName, (vbext_ProcKind)procKind);
        }

        public int GetProcCountLines(string procName, ProcKind procKind)
        {
            return InvokeResult((n, k) => ComObject.get_ProcCountLines(n, k), procName, (vbext_ProcKind)procKind);
        }

        public int GetProcBodyStartLine(string procName, ProcKind procKind)
        {
            return InvokeResult((n, k) => ComObject.get_ProcBodyLine(n, k), procName, (vbext_ProcKind)procKind);
        }

        public string GetProcOfLine(int line)
        {
            return InvokeResult(i =>
            {
                vbext_ProcKind procKind;
                return ComObject.get_ProcOfLine(line, out procKind);
            }, line);
        }

        public ProcKind GetProcKindOfLine(int line)
        {
            return InvokeResult(i =>
            {
                vbext_ProcKind procKind;
                ComObject.get_ProcOfLine(line, out procKind);
                return (ProcKind)procKind;
            }, line);
        }

        public int CountOfDeclarationLines { get { return InvokeResult(() => ComObject.CountOfDeclarationLines); } }
        public CodePane CodePane { get { return new CodePane(InvokeResult(() => ComObject.CodePane)); } }
    }
}