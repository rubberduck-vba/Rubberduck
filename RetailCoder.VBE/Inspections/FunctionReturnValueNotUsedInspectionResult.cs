using Antlr4.Runtime;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections
{
    public class FunctionReturnValueNotUsedInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public FunctionReturnValueNotUsedInspectionResult(
            IInspection inspection,
            ParserRuleContext context,
            QualifiedMemberName qualifiedName,
            IEnumerable<string> returnStatements)
            : this(inspection, context, qualifiedName, returnStatements, new List<Tuple<ParserRuleContext, QualifiedSelection, IEnumerable<string>>>())
        {
        }

        public FunctionReturnValueNotUsedInspectionResult(
            IInspection inspection,
            ParserRuleContext context,
            QualifiedMemberName qualifiedName,
            IEnumerable<string> returnStatements,
            IEnumerable<Tuple<ParserRuleContext, QualifiedSelection, IEnumerable<string>>> children)
            : base(inspection, qualifiedName.QualifiedModuleName, context)
        {
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
                // bug NullReferenceException thrown here - null Target
                return string.Format(InspectionsUI.FunctionReturnValueNotUsedInspectionResultFormat, Target.IdentifierName);
            }
        }
    }
}
