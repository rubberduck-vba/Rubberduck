using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ObsoleteTypeHintInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ObsoleteTypeHintInspectionResult(IInspection inspection, string result, QualifiedContext qualifiedContext, Declaration declaration)
            : base(inspection, result, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new RemoveTypeHintsQuickFix(Context, QualifiedSelection, declaration), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class RemoveTypeHintsQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _declaration;

        public RemoveTypeHintsQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration declaration)
            : base(context, selection, RubberduckUI.Inspections_RemoveTypeHints)
        {
            _declaration = declaration;
        }

        public override void Fix()
        {
            string hint;
            if (_declaration.HasTypeHint(out hint))
            {
                var module = _declaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
                FixTypeHintUsage(hint, module, _declaration.Selection, true);
            }

            foreach (var reference in _declaration.References)
            {
                // or should we assume type hint is the same as declaration?
                string referenceHint;
                if (reference.HasTypeHint(out referenceHint))
                {
                    var module = reference.QualifiedModuleName.Component.CodeModule;
                    FixTypeHintUsage(referenceHint, module, reference.Selection);
                }
            }

        }

        private static readonly IDictionary<string, string> TypeHints = new Dictionary<string, string>
        {
            { "%", Tokens.Integer },
            { "&", Tokens.Long },
            { "@", Tokens.Decimal },
            { "!", Tokens.Single },
            { "#", Tokens.Double },
            { "$", Tokens.String }
        };

        private void FixTypeHintUsage(string hint, CodeModule module, Selection selection, bool isDeclaration = false)
        {
            var line = module.get_Lines(selection.StartLine, 1);

            var asTypeClause = ' ' + Tokens.As + ' ' + TypeHints[hint];
            var pattern = "\\b" + _declaration.IdentifierName + "\\" + hint;
            var fix = Regex.Replace(line, pattern, _declaration.IdentifierName + (isDeclaration ? asTypeClause : String.Empty));

            module.ReplaceLine(selection.StartLine, fix);
        }
    }
}