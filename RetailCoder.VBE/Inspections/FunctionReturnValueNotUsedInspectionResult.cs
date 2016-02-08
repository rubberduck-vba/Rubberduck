using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using System.Text.RegularExpressions;
using System.Linq;
using System;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections
{
    public class FunctionReturnValueNotUsedInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public FunctionReturnValueNotUsedInspectionResult(IInspection inspection, ParserRuleContext context, QualifiedMemberName qualifiedName, IEnumerable<string> returnStatements)
            : base(inspection, string.Format(inspection.Description, qualifiedName.MemberName), qualifiedName.QualifiedModuleName, context)
        {
            _quickFixes = new[]
            {
                new ConvertToProcedureQuickFix(context, QualifiedSelection, returnStatements),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }
}
