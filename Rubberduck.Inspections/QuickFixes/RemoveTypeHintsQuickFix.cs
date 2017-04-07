using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveTypeHintsQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ObsoleteTypeHintInspection)
        };

        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        public void Fix(IInspectionResult result)
        {
            if (!string.IsNullOrWhiteSpace(result.Target.TypeHint))
            {
                var module = result.Target.QualifiedName.QualifiedModuleName.Component.CodeModule;
                {
                    FixTypeHintUsage(result.Target.TypeHint, module, result.Target.Selection, result.Target.Context, result.Target.IdentifierName, true);
                }
            }

            foreach (var reference in result.Target.References)
            {
                // or should we assume type hint is the same as declaration?
                if (!string.IsNullOrWhiteSpace(result.Target.TypeHint))
                {
                    var module = reference.QualifiedModuleName.Component.CodeModule;
                    FixTypeHintUsage(result.Target.TypeHint, module, reference.Selection, reference.Context, result.Target.IdentifierName);
                }
            }

        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveTypeHintsQuickFix;
        }

        private void FixTypeHintUsage(string hint, ICodeModule module, Selection selection, ParserRuleContext context, string identifier, bool isDeclaration = false)
        {
            var line = module.GetLines(selection.StartLine, 1);

            var asTypeClause = ' ' + Tokens.As + ' ' + SymbolList.TypeHintToTypeName[hint];

            string fix;

            if (isDeclaration && context is VBAParser.FunctionStmtContext)
            {
                var typeHint = Identifier.GetTypeHintContext(((VBAParser.FunctionStmtContext)context).functionName().identifier());
                var argList = ((VBAParser.FunctionStmtContext)context).argList();
                var endLine = argList.Stop.Line;
                var endColumn = argList.Stop.Column;

                var oldLine = module.GetLines(endLine, selection.LineCount);
                fix = oldLine.Insert(endColumn + 1, asTypeClause).Remove(typeHint.Start.Column, 1);  // adjust for VBA 0-based indexing

                module.ReplaceLine(endLine, fix);
            }
            else if (isDeclaration && context is VBAParser.PropertyGetStmtContext)
            {
                var typeHint = Identifier.GetTypeHintContext(((VBAParser.PropertyGetStmtContext)context).functionName().identifier());
                var argList = ((VBAParser.PropertyGetStmtContext)context).argList();
                var endLine = argList.Stop.Line;
                var endColumn = argList.Stop.Column;

                var oldLine = module.GetLines(endLine, selection.LineCount);
                fix = oldLine.Insert(endColumn + 1, asTypeClause).Remove(typeHint.Start.Column, 1);  // adjust for VBA 0-based indexing

                module.ReplaceLine(endLine, fix);
            }
            else
            {
                var pattern = "\\b" + identifier + "\\" + hint;
                fix = Regex.Replace(line, pattern, identifier + (isDeclaration ? asTypeClause : string.Empty));
                module.ReplaceLine(selection.StartLine, fix);
            }
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}