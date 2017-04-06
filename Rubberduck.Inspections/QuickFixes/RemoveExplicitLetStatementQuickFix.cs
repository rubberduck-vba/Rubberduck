using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveExplicitLetStatementQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ObsoleteLetStatementInspection)
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
            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            {
                if (module.IsWrappingNullReference)
                {
                    return;
                }

                var selection = result.Context.GetSelection();
                var context = (VBAParser.LetStmtContext)result.Context;

                // remove line continuations to compare against context:
                var originalCodeLines = module.GetLines(selection.StartLine, selection.LineCount)
                    .Replace("\r\n", " ")
                    .Replace("_", string.Empty);
                var originalInstruction = result.Context.GetText();

                var identifier = context.lExpression().GetText();
                var value = context.expression().GetText();

                module.DeleteLines(selection.StartLine, selection.LineCount);

                var newInstruction = identifier + " = " + value;
                var newCodeLines = originalCodeLines.Replace(originalInstruction, newInstruction);

                module.InsertLines(selection.StartLine, newCodeLines);
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveObsoleteStatementQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}