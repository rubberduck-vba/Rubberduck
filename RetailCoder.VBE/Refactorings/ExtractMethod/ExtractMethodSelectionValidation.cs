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
        private IEnumerable<Declaration> _declarations;

        public ExtractMethodSelectionValidation(IEnumerable<Declaration> declarations)
        {
            _declarations = declarations;
        }
        public bool withinSingleProcedure(QualifiedSelection qualifiedSelection)
        {

            var selection = qualifiedSelection.Selection;


            return false;
        }
    }
}
