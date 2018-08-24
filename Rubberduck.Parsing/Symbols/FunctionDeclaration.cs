using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class FunctionDeclaration : ModuleBodyElementDeclaration
    {
        public FunctionDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            Accessibility accessibility,
            ParserRuleContext context,
            ParserRuleContext attributesPassContext,
            Selection selection,
            bool isArray,
            bool isUserDefined,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(
                name,
                parent,
                parentScope,
                asTypeName,
                asTypeContext,
                typeHint,
                accessibility,
                DeclarationType.Function,
                context,
                attributesPassContext,
                selection,
                isArray,               
                isUserDefined,
                annotations,
                attributes)
        { }

        public FunctionDeclaration(ComMember member, Declaration parent, QualifiedModuleName module, Attributes attributes) 
            : this(
                module.QualifyMemberName(member.Name),
                parent,
                parent,
                member.AsTypeName.TypeName,
                null,
                null,
                Accessibility.Global,
                null,
                null,
                Selection.Home,
                member.AsTypeName.IsArray,
                false,
                null,
                attributes)
        {
            AddParameters(member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module)));
        }

        protected override bool Implements(ICanBeInterfaceMember interfaceMember)
        {
            return interfaceMember.DeclarationType == DeclarationType.Function
                && interfaceMember.IsInterfaceMember
                && IsInterfaceImplementation
                && IdentifierName.Equals($"{interfaceMember.InterfaceDeclaration.IdentifierName}_{interfaceMember.IdentifierName}");
        }
    }
}
