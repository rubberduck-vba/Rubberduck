using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MemberNotOnInterfaceInspection : InspectionBase
    {
        public MemberNotOnInterfaceInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var unresolved = State.DeclarationFinder.UnresolvedMemberDeclarations.Where(decl => !IsIgnoringInspectionResultFor(decl, AnnotationName)).ToList();

            var targets = Declarations.Where(decl => decl.AsTypeDeclaration != null &&
                                                     !decl.AsTypeDeclaration.IsUserDefined &&
                                                     decl.AsTypeDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule) &&                                                    
                                                     ((ClassModuleDeclaration)decl.AsTypeDeclaration).IsExtensible)
                                       .SelectMany(decl => decl.References).ToList();
            return unresolved
                .Select(access => new
                {
                    access,
                    callingContext = targets.FirstOrDefault(usage => usage.Context.Equals(access.CallingContext)
                                                                     || (access.CallingContext is VBAParser.NewExprContext && 
                                                                         usage.Context.Parent.Parent.Equals(access.CallingContext))
                                                                     )
                })
                .Where(memberAccess => memberAccess.callingContext != null &&
                                       memberAccess.callingContext.Declaration.DeclarationType != DeclarationType.Control)    //TODO - remove this exception after resolving #2592)
                .Select(memberAccess => new DeclarationInspectionResult(this,
                    string.Format(InspectionResults.MemberNotOnInterfaceInspection, memberAccess.access.IdentifierName,
                        memberAccess.callingContext.Declaration.AsTypeDeclaration.IdentifierName), memberAccess.access));
        }
    }
}
