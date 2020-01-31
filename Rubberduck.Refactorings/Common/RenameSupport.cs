using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    public class RenameSupport
    {
        public static IEnumerable<IGrouping<QualifiedModuleName,IdentifierReference>> RenameReferencesByQualifiedModuleName(IEnumerable<IdentifierReference> references)
        {
            var modules = references
                .Where(reference =>
                    reference.Context.GetText() != Tokens.Me
                    && !reference.IsArrayAccess
                    && !reference.IsDefaultMemberAccess)
                .GroupBy(r => r.QualifiedModuleName);

            return modules;
        }
    }
}
