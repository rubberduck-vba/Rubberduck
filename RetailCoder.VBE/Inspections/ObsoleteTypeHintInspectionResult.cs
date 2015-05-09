using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ObsoleteTypeHintInspectionResult : CodeInspectionResultBase
    {
        private readonly Declaration _declaration;

        public ObsoleteTypeHintInspectionResult(string inspection, CodeInspectionSeverity type,
            QualifiedContext qualifiedContext, Declaration declaration)
            : base(inspection, type, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _declaration = declaration;
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>
            {
                { "Remove type hints", RemoveTypeHints }
            };
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

        private void RemoveTypeHints(VBE vbe)
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

        private void FixTypeHintUsage(string hint, CodeModule module, Selection selection, bool isDeclaration = false)
        {
            var line = module.get_Lines(selection.StartLine, 1);

            var asTypeClause = ' ' + Tokens.As + ' ' + TypeHints[hint];
            var pattern = "\\b" + _declaration.IdentifierName + "\\" + hint;
            var fix = Regex.Replace(line, pattern, _declaration.IdentifierName + (isDeclaration ? asTypeClause : string.Empty));

            module.ReplaceLine(selection.StartLine, fix);
        }
    }
}