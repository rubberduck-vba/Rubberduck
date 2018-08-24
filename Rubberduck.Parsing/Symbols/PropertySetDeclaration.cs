using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class PropertySetDeclaration : PropertyDeclaration
    {
        public PropertySetDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            Accessibility accessibility,
            ParserRuleContext context,
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
                DeclarationType.PropertySet,
                context,
                selection,
                false,
                isUserDefined,
                annotations,
                attributes)
        { }

        public PropertySetDeclaration(ComMember member, Declaration parent, QualifiedModuleName module, Attributes attributes) 
            : this(
                module.QualifyMemberName(member.Name),
                parent,
                parent,
                member.AsTypeName.TypeName,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                null,
                attributes)
        {
            AddParameters(member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module)));
        }

        public PropertySetDeclaration(ComField field, Declaration parent, QualifiedModuleName module, Attributes attributes) 
            : this(
                module.QualifyMemberName(field.Name),
                parent,
                parent,
                field.ValueType,
                Accessibility.Global,
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
                   && (member.DeclarationType == DeclarationType.PropertySet
                       || member.DeclarationType == DeclarationType.Variable 
                       && (member.IsObject || member.AsTypeName.Equals(Tokens.Variant)));
        }
    }
}
