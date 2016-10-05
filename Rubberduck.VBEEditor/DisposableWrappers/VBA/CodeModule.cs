using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class CodeModule : SafeComWrapper<Microsoft.Vbe.Interop.CodeModule>, IEquatable<CodeModule>
    {
        public CodeModule(Microsoft.Vbe.Interop.CodeModule comObject) 
            : base(comObject)
        {
        }

        public VBE VBE
        {
            get { return IsWrappingNullReference ? null : new VBE(InvokeResult(() => ComObject.VBE)); }
        }

        public VBComponent Parent
        {
            get { return IsWrappingNullReference ? null : new VBComponent(InvokeResult(() => ComObject.Parent)); }
        }

        public CodePane CodePane
        {
            get { return IsWrappingNullReference ? null : new CodePane(InvokeResult(() => ComObject.CodePane)); }
        }

        public int CountOfDeclarationLines
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.CountOfDeclarationLines); }
        }

        public int CountOfLines
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.CountOfLines); }
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Name); }
            set { Invoke(() => ComObject.Name = value); }
        }

        public string GetLines(int startLine, int count)
        {
            return InvokeResult(() => ComObject.get_Lines(startLine, count));
        }

        public string Content()
        {
            return GetLines(1, CountOfLines);
        }

        private string _previousContentHash;
        public string ContentHash()
        {
            using (var hash = new SHA256Managed())
            using (var stream = Content().ToStream())
            {
                return _previousContentHash = new string(Encoding.Unicode.GetChars(hash.ComputeHash(stream)));
            }
        }

        public bool IsDirty { get { return _previousContentHash.Equals(ContentHash()); } }

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

        public int GetProcStartLine(string procName, ProcKind procKind)
        {
            return InvokeResult(() => ComObject.get_ProcStartLine(procName, (vbext_ProcKind)procKind));
        }

        public int GetProcBodyStartLine(string procName, ProcKind procKind)
        {
            return InvokeResult(() => ComObject.get_ProcBodyLine(procName, (vbext_ProcKind)procKind));
        }

        public int GetProcCountLines(string procName, ProcKind procKind)
        {
            return InvokeResult(() => ComObject.get_ProcCountLines(procName, (vbext_ProcKind)procKind));
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

        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                //CodePane.Release(); // VBE.CodePanes collection should release this CodePane
                Marshal.ReleaseComObject(ComObject);
            }
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.CodeModule> other)
        {
            return IsEqualIfNull(other) || ReferenceEquals(other.ComObject, ComObject);
        }

        public bool Equals(CodeModule other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.CodeModule>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComObject.GetHashCode();
        }
    }
}