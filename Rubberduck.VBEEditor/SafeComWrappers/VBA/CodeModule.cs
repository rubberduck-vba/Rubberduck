using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    [SuppressMessage("ReSharper", "UseIndexedProperty")]
    public class CodeModule : SafeComWrapper<VB.CodeModule>, ICodeModule
    {
        public CodeModule(VB.CodeModule target) 
            : base(target)
        {
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }

        public IVBComponent Parent
        {
            get { return new VBComponent(IsWrappingNullReference ? null : Target.Parent); }
        }

        public ICodePane CodePane
        {
            get { return new CodePane(IsWrappingNullReference ? null : Target.CodePane); }
        }

        public int CountOfDeclarationLines
        {
            get { return IsWrappingNullReference ? 0 : Target.CountOfDeclarationLines; }
        }

        public int CountOfLines
        {
            get { return IsWrappingNullReference ? 0 : Target.CountOfLines; }
        }

        public string Name
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Name; }
            set { if (!IsWrappingNullReference) Target.Name = value; }
        }

        public string GetLines(int startLine, int count)
        {
            return IsWrappingNullReference ? string.Empty : Target.get_Lines(startLine, count);
        }

        /// <summary>
        /// Returns the lines containing the selection.
        /// </summary>
        /// <param name="selection"></param>
        /// <returns></returns>
        public string GetLines(Selection selection)
        {
            return IsWrappingNullReference ? string.Empty : GetLines(selection.StartLine, selection.LineCount);
        }

        /// <summary>
        /// Deletes the lines containing the selection.
        /// </summary>
        /// <param name="selection"></param>
        public void DeleteLines(Selection selection)
        {
            if (IsWrappingNullReference) return; 
            DeleteLines(selection.StartLine, selection.LineCount);
        }

        public QualifiedSelection? GetQualifiedSelection()
        {
            if (IsWrappingNullReference || CodePane.IsWrappingNullReference)
            {
                return null;
            }
            return CodePane.GetQualifiedSelection();
        }

        public string Content()
        {
            return IsWrappingNullReference || Target.CountOfLines == 0 ? string.Empty : GetLines(1, CountOfLines);
        }

        public void Clear()
        {
            if (IsWrappingNullReference) return; 
            if (Target.CountOfLines > 0)
            {
                Target.DeleteLines(1, CountOfLines);
            }
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

        public void AddFromString(string content)
        {
            if (IsWrappingNullReference) return; 
            Target.AddFromString(content);
        }

        public void AddFromFile(string path)
        {
            if (IsWrappingNullReference) return; 
            Target.AddFromFile(path);
        }

        public void InsertLines(int line, string content)
        {
            if (IsWrappingNullReference) return; 
            Target.InsertLines(line, content);
        }

        public void DeleteLines(int startLine, int count = 1)
        {
            if (IsWrappingNullReference) return; 
            Target.DeleteLines(startLine, count);
        }

        public void ReplaceLine(int line, string content)
        {
            if (IsWrappingNullReference) return;           
            Target.ReplaceLine(line, content);
        }

        public Selection? Find(string target, bool wholeWord = false, bool matchCase = false, bool patternSearch = false)
        {
            if (IsWrappingNullReference) return null;

            var startLine = 0;
            var startColumn = 0;
            var endLine = 0;
            var endColumn = 0;

            return Target.Find(target, ref startLine, ref startColumn, ref endLine, ref endColumn, wholeWord, matchCase, patternSearch)
                ? new Selection(startLine, startColumn, endLine, endColumn)
                : (Selection?)null;
        }

        public int GetProcStartLine(string procName, ProcKind procKind)
        {
            return IsWrappingNullReference ? 0 : Target.get_ProcStartLine(procName, (vbext_ProcKind)procKind);
        }

        public int GetProcBodyStartLine(string procName, ProcKind procKind)
        {
            return IsWrappingNullReference ? 0 : Target.get_ProcBodyLine(procName, (vbext_ProcKind)procKind);
        }

        public int GetProcCountLines(string procName, ProcKind procKind)
        {
            return IsWrappingNullReference ? 0 : Target.get_ProcCountLines(procName, (vbext_ProcKind)procKind);
        }

        public string GetProcOfLine(int line)
        {
            if (IsWrappingNullReference) return string.Empty;
            vbext_ProcKind procKind;
            return Target.get_ProcOfLine(line, out procKind);
        }

        public ProcKind GetProcKindOfLine(int line)
        {
            if (IsWrappingNullReference) return 0;
            vbext_ProcKind procKind;
            Target.get_ProcOfLine(line, out procKind);
            return (ProcKind)procKind;
        }

        public override bool Equals(ISafeComWrapper<VB.CodeModule> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICodeModule other)
        {
            return Equals(other as SafeComWrapper<VB.CodeModule>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}