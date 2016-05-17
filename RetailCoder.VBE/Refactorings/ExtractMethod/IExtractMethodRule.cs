using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodRule
    {
        void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection);
    }

    public enum ExtractMethodRuleFlags
    {
        UsedBefore = 1,
        UsedAfter = 2,
        IsAssigned = 4,
        InSelection = 8

    }

    public class ExtractMethodRuleUsedBefore : IExtractMethodRule
    {
        public void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection)
        {
            if (reference.Selection.StartLine < selection.StartLine)
                flags = (byte)(flags | (byte)ExtractMethodRuleFlags.UsedBefore);
        }
    }
    public class ExtractMethodRuleUsedAfter : IExtractMethodRule
    {
        public void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection)
        {
            if (reference.Selection.StartLine > selection.EndLine)
                flags = (byte)(flags | (byte)ExtractMethodRuleFlags.UsedAfter);
        }
    }
    public class ExtractMethodRuleIsAssignedInSelection : IExtractMethodRule
    {
        public void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection)
        {
            if (selection.StartLine <= reference.Selection.StartLine && reference.Selection.StartLine <= selection.EndLine)
            {
                if (reference.IsAssignment)
                    flags = (byte)(flags | ((byte)ExtractMethodRuleFlags.IsAssigned));
            }
        }
    }
    public class ExtractMethodRuleInSelection : IExtractMethodRule
    {
        public void setValidFlag(ref byte flags, IdentifierReference reference, Selection selection)
        {
            if (selection.StartLine <= reference.Selection.StartLine && reference.Selection.StartLine <= selection.EndLine)
                flags = (byte)(flags | 8);
        }
    }

}
