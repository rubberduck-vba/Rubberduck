using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class IntegerDataTypeInspection : InspectionBase
    {
        public IntegerDataTypeInspection(RubberduckParserState state) : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var interfaceImplementationMembers = UserDeclarations.FindInterfaceImplementationMembers().ToHashSet();

            var excludeParameterMembers = State.DeclarationFinder.FindEventHandlers().ToHashSet();
            excludeParameterMembers.UnionWith(interfaceImplementationMembers);

            var result = UserDeclarations
                .Where(declaration =>
                    declaration.AsTypeName == Tokens.Integer &&
                    !interfaceImplementationMembers.Contains(declaration) &&
                    declaration.DeclarationType != DeclarationType.LibraryFunction &&
                    (declaration.DeclarationType != DeclarationType.Parameter || IncludeParameterDeclaration(declaration, excludeParameterMembers)))
                .Select(issue =>
                    new DeclarationInspectionResult(this,
                        string.Format(InspectionsUI.IntegerDataTypeInspectionResultFormat,
                            RubberduckUI.ResourceManager.GetString($"DeclarationType_{issue.DeclarationType}", CultureInfo.CurrentUICulture), issue.IdentifierName),
                        issue));

            return result;
        }

        private static bool IncludeParameterDeclaration(Declaration parameterDeclaration, ICollection<Declaration> parentDeclarationsToExclude)
        {
            var parentDeclaration = parameterDeclaration.ParentDeclaration;

            return parentDeclaration.DeclarationType != DeclarationType.LibraryFunction &&
                   parentDeclaration.DeclarationType != DeclarationType.LibraryProcedure &&
                   !parentDeclarationsToExclude.Contains(parentDeclaration);
        }
    }
}
