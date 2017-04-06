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
    public class RemoveExplicitCallStatmentQuickFix : IQuickFix
    {
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ObsoleteCallStatementInspection)
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

            var selection = result.Context.GetSelection();
            var originalCodeLines = module.GetLines(selection.StartLine, selection.LineCount);
            var originalInstruction = result.Context.GetText();

            var context = (VBAParser.CallStmtContext)result.Context;

            string target;
            string arguments;
            // The CALL statement only has arguments if it's an index expression.
            if (context.expression() is VBAParser.LExprContext && ((VBAParser.LExprContext)context.expression()).lExpression() is VBAParser.IndexExprContext)
            {
                var indexExpr = (VBAParser.IndexExprContext)((VBAParser.LExprContext)context.expression()).lExpression();
                target = indexExpr.lExpression().GetText();
                arguments = " " + indexExpr.argumentList().GetText();
            }
            else
            {
                target = context.expression().GetText();
                arguments = string.Empty;
            }
            module.DeleteLines(selection.StartLine, selection.LineCount);
            var newInstruction = target + arguments;
            var newCodeLines = originalCodeLines.Replace(originalInstruction, newInstruction);

            module.InsertLines(selection.StartLine, newCodeLines);
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
