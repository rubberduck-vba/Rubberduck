using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class SubroutineDeclaration : Declaration, IParameterizedDeclaration, ICanBeDefaultMember
    {
        private readonly List<Declaration> _parameters;

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
                  false,
                  false,
                  accessibility,
                  DeclarationType.Procedure,
                  context,
                  selection,
                  false,
                  null,
                  isUserDefined,
                  annotations,
                  attributes)
        {
            _parameters = new List<Declaration>();
        }

        public SubroutineDeclaration(ComMember member, Declaration parent, QualifiedModuleName module,
            Attributes attributes, bool eventHandler)
            : base(
                  module.QualifyMemberName(member.Name),
                  parent,
                  parent,
                  string.Empty,
                  null,
                  false,
                  false,
                  Accessibility.Global,
                  eventHandler ? DeclarationType.Event : DeclarationType.Procedure,
                  null,
                  Selection.Home,
                  false,
                  null,
                  false,
                  null,
                  attributes)
        {
            _parameters =
                member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module))
                    .Cast<Declaration>()
                    .ToList();          
        }

        public IEnumerable<Declaration> Parameters => _parameters.ToList();

        public void AddParameter(Declaration parameter)
        {
            _parameters.Add(parameter);
        }

        /// <summary>
        /// Gets an attribute value indicating whether a member is a class' default member.
        /// If this value is true, any reference to an instance of the class it's the default member of,
        /// should count as a member call to this member.
        /// </summary>
        public bool IsDefaultMember
        {
            get
            {
                IEnumerable<string> value;
                if (Attributes.TryGetValue(IdentifierName + ".VB_UserMemId", out value))
                {
                    return value.Single() == "0";
                }

                return false;
            }
        }
    }
}
