using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Inspections.Concrete
{
    [RequiredLibrary("Excel")]
    public class ExcelUdfNameIsValidCellReferenceInspection : InspectionBase
    {
        public ExcelUdfNameIsValidCellReferenceInspection(RubberduckParserState state) : base(state) { }

        private static readonly Regex ValidCellIdRegex =
            new Regex(@"^([a-z]|[a-z]{2}|[a-w][a-z]{2}|x([a-e][a-z]|f[a-d]))(?<Row>\d+)$",
                RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);

        private static readonly HashSet<Accessibility> VisibleAsUdf = new HashSet<Accessibility> { Accessibility.Public, Accessibility.Implicit };

        private const uint MaximumExcelRows = 1048576;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null)
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            var candidates = UserDeclarations.OfType<FunctionDeclaration>().Where(decl =>
                decl.ParentScopeDeclaration.DeclarationType == DeclarationType.ProceduralModule &&
                VisibleAsUdf.Contains(decl.Accessibility));

            return (from function in candidates.Where(decl => ValidCellIdRegex.IsMatch(decl.IdentifierName))
                    let row = Convert.ToUInt32(ValidCellIdRegex.Matches(function.IdentifierName)[0].Groups["Row"].Value)
                    where row > 0 && row <= MaximumExcelRows && !IsIgnoringInspectionResultFor(function, AnnotationName)
                    select new DeclarationInspectionResult(this,
                        string.Format(InspectionResults.ExcelUdfNameIsValidCellReferenceInspection, function.IdentifierName),
                        function))
                .Cast<IInspectionResult>().ToList();
        }
    }
}
