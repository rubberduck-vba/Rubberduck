using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using System.Text;
using Microsoft.VB6.Interop.VBIDE;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
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
            set { Target.Name = value; }
        }

        public string GetLines(int startLine, int count)
        {
            return Target.get_Lines(startLine, count);
        }

        /// <summary>
        /// Returns the lines containing the selection.
        /// </summary>
        /// <param name="selection"></param>
        /// <returns></returns>
        public string GetLines(Selection selection)
        {
            return GetLines(selection.StartLine, selection.LineCount);
        }

        /// <summary>
        /// Deletes the lines containing the selection.
        /// </summary>
        /// <param name="selection"></param>
        public void DeleteLines(Selection selection)
        {
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
            return Target.CountOfLines == 0 ? string.Empty : GetLines(1, CountOfLines);
        }

        public void Clear()
        {
            if (Target.CountOfLines > 0)
            {
                Target.DeleteLines(1, CountOfLines);
            }
        }

        /// <summary>
        /// Gets an array of strings where each element is a line of code in the Module,
        /// with line numbers stripped and any other pre-processing that needs to be done.
        /// </summary>
        public string[] GetSanitizedCode()
        {
            var lines = CountOfLines;
            if (lines == 0)
            {
                return new string[] { };
            }

            var code = GetLines(1, lines).Replace("\r", string.Empty).Split('\n');

            StripLineNumbers(code);
            return code;
        }

        private void StripLineNumbers(string[] lines)
        {
            var continuing = false;
            for (var line = 0; line < lines.Length; line++)
            {
                var code = lines[line];
                int? lineNumber;
                if (!continuing && HasNumberedLine(code, out lineNumber))
                {
                    var lineNumberLength = lineNumber.ToString().Length;
                    if (lines[line].Length > lineNumberLength)
                    {
                        // replace line number with as many spaces as characters taken, to avoid shifting the tokens
                        lines[line] = new string(' ', lineNumberLength) + code.Substring(lineNumber.ToString().Length + 1);
                    }
                }

                continuing = code.EndsWith(" _");
            }
        }

        private bool HasNumberedLine(string codeLine, out int? lineNumber)
        {
            lineNumber = null;

            if (string.IsNullOrWhiteSpace(codeLine.Trim()))
            {
                return false;
            }

            int line;
            var firstToken = codeLine.TrimStart().Split(' ')[0];
            if (int.TryParse(firstToken, out line))
            {
                lineNumber = line;
                return true;
            }

            return false;
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
            Target.AddFromString(content);
        }

        public void AddFromFile(string path)
        {
            Target.AddFromFile(path);
        }

        public void InsertLines(int line, string content)
        {
            Target.InsertLines(line, content);
        }

        public void DeleteLines(int startLine, int count = 1)
        {
            Target.DeleteLines(startLine, count);
        }

        public void ReplaceLine(int line, string content)
        {
            Target.ReplaceLine(line, content);
        }

        public Selection? Find(string target, bool wholeWord = false, bool matchCase = false, bool patternSearch = false)
        {
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
            return Target.get_ProcStartLine(procName, (vbext_ProcKind)procKind);
        }

        public int GetProcBodyStartLine(string procName, ProcKind procKind)
        {
            return Target.get_ProcBodyLine(procName, (vbext_ProcKind)procKind);
        }

        public int GetProcCountLines(string procName, ProcKind procKind)
        {
            return Target.get_ProcCountLines(procName, (vbext_ProcKind)procKind);
        }

        public string GetProcOfLine(int line)
        {
            vbext_ProcKind procKind;
            return Target.get_ProcOfLine(line, out procKind);
        }

        public ProcKind GetProcKindOfLine(int line)
        {
            vbext_ProcKind procKind;
            Target.get_ProcOfLine(line, out procKind);
            return (ProcKind)procKind;
        }

        public override bool Equals(ISafeComWrapper<Microsoft.VB6.Interop.VBIDE.CodeModule> other)
        {
            return IsEqualIfNull(other) || (other != null && ReferenceEquals(other.Target, Target));
        }

        public bool Equals(ICodeModule other)
        {
            return Equals(other as SafeComWrapper<Microsoft.VB6.Interop.VBIDE.CodeModule>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : Target.GetHashCode();
        }
    }
}