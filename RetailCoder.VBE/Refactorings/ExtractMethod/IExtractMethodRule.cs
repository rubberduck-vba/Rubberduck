using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodRule
    {
        Byte setValidFlag(IdentifierReference reference, Selection selection);
    }

    public enum ExtractMethodRuleFlags
    {
        UsedBefore = 1,
        UsedAfter = 2,
        IsAssigned = 4,
        InSelection = 8,
        IsExternallyReferenced = 16

    }

    public class ExtractMethodRuleUsedBefore : IExtractMethodRule
    {
        public Byte setValidFlag(IdentifierReference reference, Selection selection)
        {
            if (reference.Selection.StartLine < selection.StartLine)
                return ((byte)ExtractMethodRuleFlags.UsedBefore);
            return 0; 
        }
    }

    public class ExtractMethodRuleExternalReference : IExtractMethodRule
    {
        public Byte setValidFlag(IdentifierReference reference, Selection selection)
        {
            var decStartLine = reference.Declaration.Selection.StartLine;
            if (reference.Selection.StartLine > selection.EndLine &&
                selection.StartLine <= decStartLine &&  decStartLine <= selection.EndLine)
            {
                return ((byte)ExtractMethodRuleFlags.IsExternallyReferenced);
            }
            return 0;
        }
    }

    public class ExtractMethodRuleUsedAfter : IExtractMethodRule
    {
        public Byte setValidFlag(IdentifierReference reference, Selection selection)
        {
            if (reference.Selection.StartLine > selection.EndLine)
                return ((byte)ExtractMethodRuleFlags.UsedAfter);
            return 0;
        }
    }

    public class ExtractMethodRuleIsAssignedInSelection : IExtractMethodRule
    {
        public Byte setValidFlag(IdentifierReference reference, Selection selection)
        {
            if (selection.StartLine <= reference.Selection.StartLine && reference.Selection.StartLine <= selection.EndLine)
            {
                if (reference.IsAssignment)
                    return ((byte)ExtractMethodRuleFlags.IsAssigned);
            }
            return 0;
        }
    }

    public class ExtractMethodRuleInSelection : IExtractMethodRule
    {
        public Byte setValidFlag(IdentifierReference reference, Selection selection)
        {
            if (selection.StartLine <= reference.Selection.StartLine &&
                reference.Selection.StartLine <= selection.EndLine &&
                ((reference.Declaration == null) ? false : reference.Declaration.Selection.StartLine != reference.Selection.StartLine))
            {
                return ((byte)ExtractMethodRuleFlags.InSelection);

            }
            return 0;
        }
    }
}
