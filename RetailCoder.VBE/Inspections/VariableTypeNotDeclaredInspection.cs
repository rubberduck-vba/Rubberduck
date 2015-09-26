using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class VariableTypeNotDeclaredInspection : IInspection
    {
        public VariableTypeNotDeclaredInspection()
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public string Name { get { return "VariableTypeNotDeclaredInspection"; } }
        public string Description { get { return RubberduckUI._TypeNotDeclared_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var issues = from item in parseResult.Declarations.Items.Where(item => !item.IsBuiltIn)
                         where (item.DeclarationType == DeclarationType.Variable
                            || item.DeclarationType == DeclarationType.Constant
                            || item.DeclarationType == DeclarationType.Parameter)
                         && !item.IsTypeSpecified()
                         select new VariableTypeNotDeclaredInspectionResult(this, string.Format(Description, item.DeclarationType, item.IdentifierName), item.Context, item.QualifiedName.QualifiedModuleName);

            return issues;
        }
    }
}