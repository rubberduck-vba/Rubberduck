using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class PropertyGetDeclaration : PropertyDeclaration
    {
        public PropertyGetDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            AsTypeClauseContext asTypeContext,
            string typeHint,
            Accessibility accessibility,
            ParserRuleContext context,
            ParserRuleContext attributesPassContext,
            Selection selection,
            bool isArray,
            bool isUserDefined,
            IEnumerable<IParseTreeAnnotation> annotations,
            Attributes attributes)
            : base(
                  name,
                  parent,
                  parentScope,
                  asTypeName,
                  asTypeContext,
                  typeHint,
                  accessibility,
                  DeclarationType.PropertyGet,
                  context,
                  attributesPassContext,
                  selection,
                  isArray,
                  isUserDefined,
                  annotations,
                  attributes)
        { }

        public PropertyGetDeclaration(ComMember member, Declaration parent, QualifiedModuleName module, Attributes attributes)
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

        public PropertyGetDeclaration(ComField field, Declaration parent, QualifiedModuleName module, Attributes attributes)
            : this(
                module.QualifyMemberName(field.Name),
                parent,
                parent,
                field.ValueType,
                null,
                null,
                Accessibility.Global,
                null,
                null,
                Selection.Home,
                false,  //TODO - check this assumption.
                false,
                null,
                attributes)
        { }

        /// <inheritdoc/>
        protected override bool Implements(IInterfaceExposable member)
        {
            if (ReferenceEquals(member, this))
            {
                return false;
            }

            return (member.DeclarationType == DeclarationType.PropertyGet || member.DeclarationType == DeclarationType.Variable)
                   && member.IsInterfaceMember
                   && ((ClassModuleDeclaration)member.ParentDeclaration).Subtypes.Any(implementation => ReferenceEquals(implementation, ParentDeclaration))
                   && IdentifierName.Equals(member.ImplementingIdentifierName);
        }

        public override BlockContext Block => ((PropertyGetStmtContext)Context).block();
    }
}
