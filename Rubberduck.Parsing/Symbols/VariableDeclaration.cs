using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class VariableDeclaration : Declaration, IInterfaceExposable
    {
        public VariableDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration parentDeclaration,
            Declaration parentScope,
            string asTypeName,
            string typeHint,
            bool isSelfAssigned,
            bool isWithEvents,
            Accessibility accessibility,
            ParserRuleContext context,
            ParserRuleContext attributesPassContext,
            Selection selection,
            bool isArray,
            VBAParser.AsTypeClauseContext asTypeContext,
            IEnumerable<IAnnotation> annotations = null,
            Attributes attributes = null)
            : base(
                qualifiedName,
                parentDeclaration,
                parentScope?.Scope,
                asTypeName,
                typeHint,
                isSelfAssigned,
                isWithEvents,
                accessibility,
                DeclarationType.Variable,
                context,
                attributesPassContext,
                selection,
                isArray,
                asTypeContext,
                true,
                annotations,
                attributes)
        {
            if ((accessibility == Accessibility.Public || accessibility == Accessibility.Implicit) 
                && parentDeclaration is ClassModuleDeclaration classModule)
            {
                classModule.AddMember(this);
            }
        }

        /// <summary>
        /// True if a variable is declared with the <c>Static</c> keyword, or declared in a procedure scope that uses this keyword.
        /// </summary>
        /// <remarks>
        /// In VBA/VB6, a Static variable keeps its value between procedure calls.
        /// </remarks>
        public bool IsStatic => ParentScopeDeclaration is ModuleBodyElementDeclaration parent && parent.IsStatic
                                || Context is VBAParser.VariableStmtContext context && context.STATIC() != null;

        /// <inheritdoc/>
        public string ImplementingIdentifierName => this.ImplementingIdentifierName();

        /// <inheritdoc/>
        public bool IsInterfaceMember => this.IsInterfaceMember();

        /// <inheritdoc/>
        public ClassModuleDeclaration InterfaceDeclaration => this.InterfaceDeclaration();
    }
}
