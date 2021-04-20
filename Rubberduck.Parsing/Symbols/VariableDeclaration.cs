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
            IEnumerable<IParseTreeAnnotation> annotations = null,
            Attributes attributes = null, 
            bool isUserDefined = true)
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
                isUserDefined,
                annotations,
                attributes)
        {
            if ((accessibility == Accessibility.Public || accessibility == Accessibility.Implicit) 
                && parentDeclaration is ClassModuleDeclaration classModule)
            {
                classModule.AddMember(this);
            }
        }

        /// <inheritdoc/>
        public string ImplementingIdentifierName => this.ImplementingIdentifierName();

        /// <inheritdoc/>
        public bool IsInterfaceMember => this.IsInterfaceMember();

        /// <inheritdoc/>
        public ClassModuleDeclaration InterfaceDeclaration => this.InterfaceDeclaration();
    }
}
