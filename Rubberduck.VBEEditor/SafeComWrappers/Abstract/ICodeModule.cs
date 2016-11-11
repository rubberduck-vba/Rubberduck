using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ICodeModule : ISafeComWrapper, IEquatable<ICodeModule>
    {
        IVBE VBE { get; }
        IVBComponent Parent { get; }
        ICodePane CodePane { get; }
        int CountOfDeclarationLines { get; }
        int CountOfLines { get; }
        string Name { get; set; }
        string GetLines(int startLine, int count);
        string GetLines(Selection selection);
        void DeleteLines(Selection selection);
        void DeleteLines(int startLine, int count = 1);
        QualifiedSelection? GetQualifiedSelection();
        string Content();
        void Clear();
        string ContentHash();
        void AddFromString(string content);
        void AddFromFile(string path);
        void InsertLines(int line, string content);
        void ReplaceLine(int line, string content);
        int GetProcStartLine(string procName, ProcKind procKind);
        int GetProcBodyStartLine(string procName, ProcKind procKind);
        int GetProcCountLines(string procName, ProcKind procKind);
        string GetProcOfLine(int line);
        ProcKind GetProcKindOfLine(int line);
    }
}