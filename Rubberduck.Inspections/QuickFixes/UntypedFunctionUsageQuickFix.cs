using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class UntypedFunctionUsageQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(UntypedFunctionUsageInspection)
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
            var originalInstruction = result.Context.GetText();
            var newInstruction = GetNewSignature(result.Context);
            var selection = result.QualifiedSelection.Selection;

            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            var lines = module.GetLines(selection.StartLine, selection.LineCount);

            var newContent = lines.Remove(result.Context.Start.Column, originalInstruction.Length)
                .Insert(result.Context.Start.Column, newInstruction);
            module.ReplaceLine(selection.StartLine, newContent);
        }

        public string Description(IInspectionResult result)
        {
            return string.Format(InspectionsUI.QuickFixUseTypedFunction_, result.Context.GetText(), GetNewSignature(result.Context));
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

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}