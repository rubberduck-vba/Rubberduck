using Antlr4.Runtime;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections
{
    public class FunctionReturnValueNotUsedInspectionResult : InspectionResultBase
    {
        private readonly string _identifierName;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public FunctionReturnValueNotUsedInspectionResult(
            IInspection inspection,
            ParserRuleContext context,
            QualifiedMemberName qualifiedName,
            IEnumerable<string> returnStatements,
            string identifierName)
            : this(inspection, context, qualifiedName, returnStatements, new List<Tuple<ParserRuleContext, QualifiedSelection, IEnumerable<string>>>(), identifierName)
        {
        }

        public FunctionReturnValueNotUsedInspectionResult(
            IInspection inspection,
            ParserRuleContext context,
            QualifiedMemberName qualifiedName,
            IEnumerable<string> returnStatements,
            IEnumerable<Tuple<ParserRuleContext, QualifiedSelection, IEnumerable<string>>> children,
            string identifierName)
            : base(inspection, qualifiedName.QualifiedModuleName, context)
        {
            _identifierName = identifierName;
            var root = new ConvertToProcedureQuickFix(context, QualifiedSelection, returnStatements);
            var compositeFix = new CompositeCodeInspectionFix(root);
            children.ToList().ForEach(child => compositeFix.AddChild(new ConvertToProcedureQuickFix(child.Item1, child.Item2, child.Item3)));
            _quickFixes = new[]
            {
                compositeFix
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.FunctionReturnValueNotUsedInspectionResultFormat, _identifierName);
            }
        }
    }
}
