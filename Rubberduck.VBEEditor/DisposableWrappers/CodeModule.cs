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
            Invoke(() => ComObject.AddFromString(value));
        }

        public void AddFromFile(string path)
        {
            Invoke(() => ComObject.AddFromFile(path));
        }

        public void InsertLines(int line, string content)
        {
            Invoke(() => ComObject.InsertLines(line, content));
        }

        public void DeleteLines(int startLine, int count = 1)
        {
            Invoke(() => ComObject.DeleteLines(startLine, count));
        }

        public void ReplaceLine(int line, string content)
        {
            Invoke(() => ComObject.ReplaceLine(line, content));
        }

        public int CreateEventProc(string eventName, string objectName)
        {
            return InvokeResult(() => ComObject.CreateEventProc(eventName, objectName));
        }

        public Selection? Find(string target, bool wholeWord = false, bool matchCase = false, bool patternSearch = false)
        {
            return InvokeResult(() =>
            {
                var startLine = 0;
                var startColumn = 0;
                var endLine = 0;
                var endColumn = 0;
                if (ComObject.Find(target, ref startLine, ref startColumn, ref endLine, ref endColumn, wholeWord, matchCase, patternSearch))
                {
                    return new Selection(startLine, startColumn, endLine, endColumn);
                }
                return (Selection?)null;
            });
        }

        public VBComponent Parent { get; private set; }
        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        public string Name
        {
            get { return InvokeResult(() => ComObject.Name); }
            set { Invoke(() => ComObject.Name = value); }
        }

        public string GetLines(int startLine, int count)
        {
            return InvokeResult(() => ComObject.get_Lines(startLine, count));
        }

        public int CountOfLines { get { return InvokeResult(() => ComObject.CountOfLines); } }
        public int GetProcStartLine(string procName, ProcKind procKind)
        {
            return InvokeResult(() => ComObject.get_ProcStartLine(procName, (vbext_ProcKind)procKind));
        }

        public int GetProcCountLines(string procName, ProcKind procKind)
        {
            return InvokeResult(() => ComObject.get_ProcCountLines(procName, (vbext_ProcKind)procKind));
        }

        public int GetProcBodyStartLine(string procName, ProcKind procKind)
        {
            return InvokeResult(() => ComObject.get_ProcBodyLine(procName, (vbext_ProcKind)procKind));
        }

        public string GetProcOfLine(int line)
        {
            return InvokeResult(() =>
            {
                vbext_ProcKind procKind;
                return ComObject.get_ProcOfLine(line, out procKind);
            });
        }

        public ProcKind GetProcKindOfLine(int line)
        {
            return InvokeResult(() =>
            {
                vbext_ProcKind procKind;
                ComObject.get_ProcOfLine(line, out procKind);
                return (ProcKind)procKind;
            });
        }

        public int CountOfDeclarationLines { get { return InvokeResult(() => ComObject.CountOfDeclarationLines); } }
        public CodePane CodePane { get { return new CodePane(InvokeResult(() => ComObject.CodePane)); } }
    }
}