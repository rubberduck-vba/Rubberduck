using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class UntypedFunctionUsageInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public UntypedFunctionUsageInspectionResult(IInspection inspection, IdentifierReference reference) 
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new UntypedFunctionUsageQuickFix((ParserRuleContext)GetFirst(typeof(VBAParser.IdentifierContext)).Parent, QualifiedSelection), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(Inspection.Description, _reference.Declaration.IdentifierName); }
        }

        private ParserRuleContext GetFirst(Type nodeType)
        {
            var unexploredNodes = new List<ParserRuleContext> {Context};

            while (unexploredNodes.Any())
            {
                if (unexploredNodes[0].GetType() == nodeType)
                {
                    return unexploredNodes[0];
                }
                
                unexploredNodes.AddRange(unexploredNodes[0].children.OfType<ParserRuleContext>());
                unexploredNodes.RemoveAt(0);
            }

            return null;
        }
    }

    public class UntypedFunctionUsageQuickFix : CodeInspectionQuickFix
    {
        public UntypedFunctionUsageQuickFix(ParserRuleContext context, QualifiedSelection selection) 
            : base(context, selection, string.Format(InspectionsUI.QuickFixUseTypedFunction_, context.GetText(), GetNewSignature(context)))
        {
        }

        public override void Fix()
        {
            var originalInstruction = Context.GetText();
            var newInstruction = GetNewSignature(Context);
            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.Lines[selection.StartLine, selection.LineCount];

            var result = lines.Replace(originalInstruction, newInstruction);
            module.ReplaceLine(selection.StartLine, result);
        }

        private static string GetNewSignature(ParserRuleContext context)
        {
            Debug.Assert(context != null);

            return context.children.Aggregate(string.Empty, (current, member) =>
            {
                var isIdentifierNode = member is VBAParser.IdentifierContext;
                return current + member.GetText() + (isIdentifierNode ? "$" : string.Empty);
            });
        }
    }
}
