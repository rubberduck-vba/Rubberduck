using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.Results
{
    public class ParameterNotUsedInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ParameterNotUsedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context, QualifiedMemberName qualifiedName, bool isInterfaceImplementation, 
            IVBE vbe, RubberduckParserState state, IMessageBox messageBox)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target)
        {
            _quickFixes = isInterfaceImplementation ? new QuickFixBase[] {} : new QuickFixBase[]
            {
                new RemoveUnusedParameterQuickFix(Context, QualifiedSelection, vbe, state, messageBox),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterNotUsedInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }
    }
}
