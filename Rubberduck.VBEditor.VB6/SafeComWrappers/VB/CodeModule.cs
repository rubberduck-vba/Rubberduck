using System;
using System.Diagnostics.CodeAnalysis;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    [SuppressMessage("ReSharper", "UseIndexedProperty")]
    public class CodeModule : SafeComWrapper<VB.CodeModule>, ICodeModule
    {
        public CodeModule(VB.CodeModule target, bool rewrapping = false) 
            : base(target, rewrapping)
        {
        }

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IVBComponent Parent => new VBComponent(IsWrappingNullReference ? null : Target.Parent);

        public ICodePane CodePane => new CodePane(IsWrappingNullReference ? null : Target.CodePane);

        public int CountOfDeclarationLines => IsWrappingNullReference ? 0 : Target.CountOfDeclarationLines;

        public int CountOfLines => IsWrappingNullReference ? 0 : Target.CountOfLines;

        public string Name
        {
            get => IsWrappingNullReference ? string.Empty : Target.Name;
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
            if (IsWrappingNullReference)
            {
                return null;
            }

            using (var codePane = CodePane)
            {
                if (codePane.IsWrappingNullReference)
                {
                    return null;
                }
            
                return codePane.GetQualifiedSelection();
            }
        }

        public QualifiedModuleName QualifiedModuleName
        {
            get
            {
                using (var component = Parent)
                {
                    return component.QualifiedModuleName;
                }
            }
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

        public int ContentHash()
        {
            var code = Content();
            return string.IsNullOrEmpty(code)
                ? 0
                : code.GetHashCode();
        }

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

            try
            {
                Target.InsertLines(line, content);
            }
            catch (Exception e)
            {
                // "too many line continuations" is one possible cause for a COMException here.
                // deleting the only line in a module is another.
                // we can log the exception, but really we're just intentionally swallowing it.
                _logger.Warn(e, $"{nameof(InsertLines)} failed.");
            }
        }

        public void DeleteLines(int startLine, int count = 1)
        {
            if (IsWrappingNullReference) return;

            try
            {
                Target.DeleteLines(startLine, count);
            }
            catch (Exception e)
            {
                // "too many line continuations" is one possible cause for a COMException here.
                // deleting the only line in a module is another.
                // we can log the exception, but really we're just intentionally swallowing it.
                _logger.Warn(e, $"{nameof(DeleteLines)} failed.");
            }
        }

        public void ReplaceLine(int line, string content)
        {
            if (IsWrappingNullReference) return;

            try
            {
                Target.ReplaceLine(line, content);
                if (Target.CountOfLines == 0)
                {
                    Target.AddFromString(content);
                }
            }
            catch (Exception e)
            {
                // "too many line continuations" is one possible cause for a COMException here.
                // deleting the only line in a module is another.
                // we can log the exception, but really we're just intentionally swallowing it.
                _logger.Warn(e, $"{nameof(ReplaceLine)} failed.");
            }
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
            return IsWrappingNullReference ? 0 : Target.get_ProcStartLine(procName, (VB.vbext_ProcKind)procKind);
        }

        public int GetProcBodyStartLine(string procName, ProcKind procKind)
        {
            return IsWrappingNullReference ? 0 : Target.get_ProcBodyLine(procName, (VB.vbext_ProcKind)procKind);
        }

        public int GetProcCountLines(string procName, ProcKind procKind)
        {
            return IsWrappingNullReference ? 0 : Target.get_ProcCountLines(procName, (VB.vbext_ProcKind)procKind);
        }

        public string GetProcOfLine(int line)
        {
            return IsWrappingNullReference ? string.Empty : Target.get_ProcOfLine(line, out _);
        }

        public ProcKind GetProcKindOfLine(int line)
        {
            if (IsWrappingNullReference) return 0;
            Target.get_ProcOfLine(line, out var procKind);
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

        protected override void Dispose(bool disposing) => base.Dispose(disposing);
    }
}