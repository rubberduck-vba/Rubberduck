using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class PropertyLetDeclaration : PropertyDeclaration
    {
        public PropertyLetDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            Accessibility accessibility,
            ParserRuleContext context,
            ParserRuleContext attributesPassContext,
            Selection selection,
            bool isUserDefined,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(
                name,
                parent,
                parentScope,
                asTypeName,
                null,
                null,
                accessibility,
                DeclarationType.PropertyLet,
                context,
                attributesPassContext,
                selection,
                false,
                isUserDefined,
                annotations,
                attributes)
        { }

        public PropertyLetDeclaration(ComMember member, Declaration parent, QualifiedModuleName module, Attributes attributes)
            : this(
                module.QualifyMemberName(member.Name),
                parent,
                parent,
                member.AsTypeName.TypeName, 
                Accessibility.Global,
                null,
                null,
                Selection.Home,
                false,
                null,
                attributes)
        {
            AddParameters(member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module)));
        }

        public PropertyLetDeclaration(ComField field, Declaration parent, QualifiedModuleName module, Attributes attributes)
            : this(
                module.QualifyMemberName(field.Name),
                parent,
                parent,
                field.ValueType,
                Accessibility.Global,
                null,
                null,
                Selection.Home,
                false,
                null,
                attributes)
        { }

        protected override bool Implements(ICanBeInterfaceMember member)
        {
            return member.IsInterfaceMember
                   && IsInterfaceImplementation
                   && IdentifierName.Equals($"{member.InterfaceDeclaration.IdentifierName}_{member.IdentifierName}")
                   && (member.DeclarationType == DeclarationType.PropertyLet
                       || member.DeclarationType == DeclarationType.Variable
                       && !member.IsObject);
        }
    }
}
