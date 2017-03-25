using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// These declarations are created from unresolved member accesses in the DeclarationFinder and are collected for use by inspections.  They
    /// should NOT be added to the Declaration collections in the parser state.
    /// </summary>
    public class UnboundMemberDeclaration : Declaration
    {
        /// <summary>
        /// Context on the LHS of the member access.
        /// </summary>
        public ParserRuleContext CallingContext { get; private set; }

        public UnboundMemberDeclaration(Declaration parentDeclaration, ParserRuleContext unboundIdentifier, ParserRuleContext callingContext, IEnumerable<IAnnotation> annotations) :
            base(new QualifiedMemberName(parentDeclaration.QualifiedName.QualifiedModuleName, unboundIdentifier.GetText()),
                 parentDeclaration,
                 parentDeclaration,
                 "Variant",
                 string.Empty,
                 false,
                 false,
                 Accessibility.Implicit, 
                 DeclarationType.UnresolvedMember, 
                 unboundIdentifier, 
                 unboundIdentifier.GetSelection(), 
                 false, 
                 null,
                 true, 
                 annotations)
        {
            CallingContext = callingContext;
        }
    }
}
