using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class MemberNotOnInterfaceInspection : InspectionBase
    {
        public MemberNotOnInterfaceInspection(RubberduckParserState state, CodeInspectionSeverity defaultSeverity = CodeInspectionSeverity.Warning)
            : base(state, defaultSeverity)
        {
        }

        public override string Meta { get { return InspectionsUI.MemberNotOnInterfaceInspectionMeta; } }
        public override string Description { get { return InspectionsUI.MemberNotOnInterfaceInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var unresolved = State.DeclarationFinder.UnresolvedMemberDeclarations().Where(decl => !IsIgnoringInspectionResultFor(decl, AnnotationName)).ToList();

            var targets = Declarations.Where(decl => decl.AsTypeDeclaration != null &&
                                                     decl.AsTypeDeclaration.IsBuiltIn &&
                                                     decl.AsTypeDeclaration.DeclarationType == DeclarationType.ClassModule &&
                                                     ((ClassModuleDeclaration)decl.AsTypeDeclaration).IsExtensible)
                                       .SelectMany(decl => decl.References).ToList();

            return (from access in unresolved
                let callingContext = targets.FirstOrDefault(usage => usage.Context.Equals(access.CallingContext))
                where callingContext != null
                select
                    new MemberNotOnInterfaceInspectionResult(this, access, callingContext.Declaration.AsTypeDeclaration))
                .ToList();
        }
    }
}
