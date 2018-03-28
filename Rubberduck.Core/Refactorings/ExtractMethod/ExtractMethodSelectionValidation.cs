using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
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
            IEnumerable<Declaration> procedures = _declarations.Where(d => d.IsUserDefined && (DeclarationExtensions.ProcedureTypes.Contains(d.DeclarationType)));
            Func<int, dynamic> ProcOfLine = (sl) => procedures.FirstOrDefault(d => d.Context.Start.Line < sl && d.Context.Stop.Line > sl);

            var startLine = selection.StartLine;
            var endLine = selection.EndLine;

            // End of line is easy
            var procEnd = ProcOfLine(endLine);
            if (procEnd == null)
            {
                return false;
            }

            var procEndContext = procEnd.Context as ParserRuleContext;
            var procEndLine = procEndContext.Stop.Line;

            /* Handle: function signature continuations
             * public function(byval a as string _
             *                 byval b as string) as integer
             */
            var procStart = ProcOfLine(startLine);
            if (procStart == null)
            {
                return false;
            }

            dynamic procStartContext;
            procStartContext = procStart.Context as VBAParser.FunctionStmtContext;
            if (procStartContext == null)
            {
                procStartContext = procStart.Context as VBAParser.SubStmtContext;
            }
            // TOOD: Doesn't support properties.
            if (procStartContext == null)
            {
                return false;
            }
            var procEndOfSignature = procStartContext.endOfStatement() as VBAParser.EndOfStatementContext;
            var procSignatureLastLine = procEndOfSignature.Start.Line;

            return (procEnd as Declaration).QualifiedSelection.Equals((procStart as Declaration).QualifiedSelection)
                && (procSignatureLastLine < startLine);

        }
    }
}
