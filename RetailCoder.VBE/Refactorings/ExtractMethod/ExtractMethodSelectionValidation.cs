using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodSelectionValidation : IExtractMethodSelectionValidation
    {
        private IEnumerable<Declaration> declarations;

        public ExtractMethodSelectionValidation(IEnumerable<Declaration> declarations)
        {
            // TODO: Complete member initialization
            this.declarations = declarations;
        }
        public bool withinSingleProcedure(QualifiedSelection qualifiedSelection)
        {
            throw new NotImplementedException();
        }
    }
}
