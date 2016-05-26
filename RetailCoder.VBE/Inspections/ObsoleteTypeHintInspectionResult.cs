using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public class ObsoleteTypeHintInspectionResult : InspectionResultBase
    {
        private readonly string _result;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ObsoleteTypeHintInspectionResult(IInspection inspection, string result, QualifiedContext qualifiedContext, Declaration declaration)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _result = result;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new RemoveTypeHintsQuickFix(Context, QualifiedSelection, declaration), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return _result; }
        }
    }

    public class RemoveTypeHintsQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _declaration;

        public RemoveTypeHintsQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration declaration)
            : base(context, selection, InspectionsUI.RemoveTypeHintsQuickFix)
        {
            _declaration = declaration;
        }

        public override void Fix()
        {
            if (!string.IsNullOrWhiteSpace(_declaration.TypeHint))
            {
                var module = _declaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
                FixTypeHintUsage(_declaration.TypeHint, module, _declaration.Selection, true);
            }

            foreach (var reference in _declaration.References)
            {
                // or should we assume type hint is the same as declaration?
                if (!string.IsNullOrWhiteSpace(_declaration.TypeHint))
                {
                    var module = reference.QualifiedModuleName.Component.CodeModule;
                    FixTypeHintUsage(_declaration.TypeHint, module, reference.Selection);
                }
            }

        }

        private void FixTypeHintUsage(string hint, CodeModule module, Selection selection, bool isDeclaration = false)
        {
            var line = module.Lines[selection.StartLine, 1];

            var asTypeClause = ' ' + Tokens.As + ' ' + Declaration.TYPEHINT_TO_TYPENAME[hint];

            string fix;

            if (isDeclaration && Context is VBAParser.FunctionStmtContext)
            {
                var typeHint = Identifier.GetTypeHintContext(((VBAParser.FunctionStmtContext)Context).functionName().identifier());
                var argList = ((VBAParser.FunctionStmtContext)Context).argList();
                var endLine = argList.Stop.Line;
                var endColumn = argList.Stop.Column;

                var oldLine = module.Lines[endLine, selection.LineCount];
                fix = oldLine.Insert(endColumn + 1, asTypeClause).Remove(typeHint.Start.Column, 1);  // adjust for VBA 0-based indexing

                module.ReplaceLine(endLine, fix);
            }
            else if (isDeclaration && Context is VBAParser.PropertyGetStmtContext)
            {
                var typeHint = Identifier.GetTypeHintContext(((VBAParser.PropertyGetStmtContext)Context).functionName().identifier());
                var argList = ((VBAParser.PropertyGetStmtContext)Context).argList();
                var endLine = argList.Stop.Line;
                var endColumn = argList.Stop.Column;

                var oldLine = module.Lines[endLine, selection.LineCount];
                fix = oldLine.Insert(endColumn + 1, asTypeClause).Remove(typeHint.Start.Column, 1);  // adjust for VBA 0-based indexing

                module.ReplaceLine(endLine, fix);
            }
            else
            {
                var pattern = "\\b" + _declaration.IdentifierName + "\\" + hint;
                fix = Regex.Replace(line, pattern, _declaration.IdentifierName + (isDeclaration ? asTypeClause : string.Empty));
                module.ReplaceLine(selection.StartLine, fix);
            }
        }
    }
}
