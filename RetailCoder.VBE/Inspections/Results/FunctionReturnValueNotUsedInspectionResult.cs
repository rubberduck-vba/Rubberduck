using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{


    public class FunctionReturnValueNotUsedInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public FunctionReturnValueNotUsedInspectionResult(
            IInspection inspection,
            ParserRuleContext context,
            QualifiedMemberName qualifiedName,
            Declaration target,
            bool allowConvertToProcedure = true)
            : this(inspection, context, qualifiedName, new List<Tuple<ParserRuleContext, QualifiedSelection, Declaration>>(), target, allowConvertToProcedure)
        {
        }

        public FunctionReturnValueNotUsedInspectionResult(
            IInspection inspection,
            ParserRuleContext context,
            QualifiedMemberName qualifiedName,
            IEnumerable<Tuple<ParserRuleContext, QualifiedSelection, Declaration>> children,
            Declaration target, bool allowConvertToProcedure = true)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target)
        {
            var ignoreOnce = new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName);
            if (allowConvertToProcedure)
            {
                var root = new ConvertToProcedureQuickFix(context, QualifiedSelection, target);
                var compositeFix = new CompositeCodeInspectionFix(root);
                children.ToList().ForEach(child => compositeFix.AddChild(new ConvertToProcedureQuickFix(child.Item1, child.Item2, child.Item3)));
                _quickFixes = new QuickFixBase[]
                {
                    compositeFix,
                    ignoreOnce
                };
            }
            else
            {
                _quickFixes = new[] {ignoreOnce};
            }
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.FunctionReturnValueNotUsedInspectionResultFormat, Target.IdentifierName).Captialize();
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
