using System.Text.RegularExpressions;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveTypeHintsQuickFix : QuickFixBase
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
                {
                    FixTypeHintUsage(_declaration.TypeHint, module, _declaration.Selection, true);
                }
            }

            foreach (var reference in _declaration.References)
            {
                // or should we assume type hint is the same as declaration?
                if (!string.IsNullOrWhiteSpace(_declaration.TypeHint))
                {
                    var module = reference.QualifiedModuleName.Component.CodeModule;
                    {
                        FixTypeHintUsage(_declaration.TypeHint, module, reference.Selection);
                    }
                }
            }

        }

        private void FixTypeHintUsage(string hint, ICodeModule module, Selection selection, bool isDeclaration = false)
        {
            var line = module.GetLines(selection.StartLine, 1);

            var asTypeClause = ' ' + Tokens.As + ' ' + SymbolList.TypeHintToTypeName[hint];

            string fix;

            if (isDeclaration && Context is VBAParser.FunctionStmtContext)
            {
                var typeHint = Identifier.GetTypeHintContext(((VBAParser.FunctionStmtContext)Context).functionName().identifier());
                var argList = ((VBAParser.FunctionStmtContext)Context).argList();
                var endLine = argList.Stop.Line;
                var endColumn = argList.Stop.Column;

                var oldLine = module.GetLines(endLine, selection.LineCount);
                fix = oldLine.Insert(endColumn + 1, asTypeClause).Remove(typeHint.Start.Column, 1);  // adjust for VBA 0-based indexing

                module.ReplaceLine(endLine, fix);
            }
            else if (isDeclaration && Context is VBAParser.PropertyGetStmtContext)
            {
                var typeHint = Identifier.GetTypeHintContext(((VBAParser.PropertyGetStmtContext)Context).functionName().identifier());
                var argList = ((VBAParser.PropertyGetStmtContext)Context).argList();
                var endLine = argList.Stop.Line;
                var endColumn = argList.Stop.Column;

                var oldLine = module.GetLines(endLine, selection.LineCount);
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