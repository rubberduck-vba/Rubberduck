using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class SplitMultipleDeclarationsQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(MultipleDeclarationsInspection)
        };

        public SplitMultipleDeclarationsQuickFix(RubberduckParserState state)
        {
            _state = state;
        }
        
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
            dynamic context = result.Context is VBAParser.ConstStmtContext
                ? result.Context
                : result.Context.Parent;

            var declarationsText = GetDeclarationsText(context);

            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Replace(context, declarationsText);
        }

        private string GetDeclarationsText(VBAParser.ConstStmtContext consts)
        {
            var keyword = string.Empty;
            if (consts.visibility() != null)
            {
                keyword += consts.visibility().GetText() + ' ';
            }

            keyword += consts.CONST().GetText() + ' ';

            var newContent = new StringBuilder();
            foreach (var constant in consts.constSubStmt())
            {
                newContent.AppendLine(keyword + constant.GetText());
            }

            return newContent.ToString();
        }

        private string GetDeclarationsText(VBAParser.VariableStmtContext variables)
        {
            var keyword = string.Empty;
            if (variables.DIM() != null)
            {
                keyword += Tokens.Dim + ' ';
            }
            else if (variables.visibility() != null)
            {
                keyword += variables.visibility().GetText() + ' ';
            }
            else if (variables.STATIC() != null)
            {
                keyword += variables.STATIC().GetText() + ' ';
            }

            var newContent = new StringBuilder();
            foreach (var variable in variables.variableListStmt().variableSubStmt())
            {
                newContent.AppendLine(keyword + variable.GetText());
            }

            return newContent.ToString();
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.SplitMultipleDeclarationsQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}