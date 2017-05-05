using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MemberNotOnInterfaceInspection : InspectionBase
    {
        public MemberNotOnInterfaceInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var unresolved = State.DeclarationFinder.UnresolvedMemberDeclarations().Where(decl => !IsIgnoringInspectionResultFor(decl, AnnotationName)).ToList();

            var targets = Declarations.Where(decl => decl.AsTypeDeclaration != null &&
                                                     !decl.AsTypeDeclaration.IsUserDefined &&
                                                     decl.AsTypeDeclaration.DeclarationType == DeclarationType.ClassModule &&
                                                     ((ClassModuleDeclaration)decl.AsTypeDeclaration).IsExtensible)
                                       .SelectMany(decl => decl.References).ToList();

            return from access in unresolved
                   let callingContext = targets.FirstOrDefault(usage => usage.Context.Equals(access.CallingContext))
                   where callingContext != null
                   select new DeclarationInspectionResult(this,
                                               string.Format(InspectionsUI.MemberNotOnInterfaceInspectionResultFormat, access.IdentifierName, callingContext.Declaration.AsTypeDeclaration.IdentifierName),
                                               access);
        }
    }
}
