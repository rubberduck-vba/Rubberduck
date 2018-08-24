using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class SubroutineDeclaration : ModuleBodyElementDeclaration
    {
        public SubroutineDeclaration(
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
                  string.Empty,
                  accessibility,
                  DeclarationType.Procedure,
                  context,
                  selection,
                  false,
                  isUserDefined,
                  annotations,
                  attributes)
        { }

        public SubroutineDeclaration(ComMember member, Declaration parent, QualifiedModuleName module, Attributes attributes, bool eventHandler)
            : base(
                  module.QualifyMemberName(member.Name),
                  parent,
                  parent,
                  string.Empty,
                  null,
                  string.Empty,
                  Accessibility.Global,
                  eventHandler ? DeclarationType.Event : DeclarationType.Procedure,
                  null,
                  Selection.Home,
                  false,
                  false,
                  null,
                  attributes)
        {
            AddParameters(member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module)));
        }

        protected override bool Implements(ICanBeInterfaceMember interfaceMember)
        {
            return DeclarationType == DeclarationType.Procedure
                   && interfaceMember.DeclarationType == DeclarationType.Procedure
                   && interfaceMember.IsInterfaceMember
                   && IsInterfaceImplementation
                   && IdentifierName.Equals($"{interfaceMember.InterfaceDeclaration.IdentifierName}_{interfaceMember.IdentifierName}");
        }
    }
}
