using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class EventDeclaration : Declaration, IParameterizedDeclaration
    {
        private readonly List<ParameterDeclaration> _parameters;

        public EventDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            Accessibility accessibility,
            ParserRuleContext context,
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
                  typeHint,
                  false,
                  false,
                  accessibility,
                  DeclarationType.Event,
                  context,
                  null,
                  selection,
                  isArray,
                  asTypeContext,
                  isUserDefined,
                  annotations,
                  attributes)
        {
            _parameters = new List<ParameterDeclaration>();
        }

        public EventDeclaration(ComMember member, Declaration parent, QualifiedModuleName module,
            Attributes attributes) : this(
                module.QualifyMemberName(member.Name),
                parent,
                parent,
                member.AsTypeName.TypeName,
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                member.AsTypeName.IsArray,
                false,
                null,
                attributes)
        {
            _parameters =
                member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module))
                    .ToList();
        }

        public IReadOnlyList<ParameterDeclaration> Parameters => _parameters.ToList();

        public void AddParameter(ParameterDeclaration parameter)
        {
            _parameters.Add(parameter);
        }
    }
}
