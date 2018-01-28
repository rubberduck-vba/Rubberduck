using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ExternalProcedureDeclaration : Declaration, IParameterizedDeclaration
    {
        private readonly List<ParameterDeclaration> _parameters;

        public ExternalProcedureDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            DeclarationType declarationType,
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            Accessibility accessibility,
            ParserRuleContext context,
            Selection selection,
            bool isUserDefined,
            IEnumerable<IAnnotation> annotations)
            : base(
                  name,
                  parent,
                  parentScope,
                  asTypeName,
                  null,
                  false,
                  false,
                  accessibility,
                  declarationType,
                  context,
                  selection,
                  false,
                  asTypeContext,
                  isUserDefined,
                  annotations,
                  null)
        {
            _parameters = new List<ParameterDeclaration>();
        }

        public IEnumerable<ParameterDeclaration> Parameters => _parameters.ToList();

        public void AddParameter(ParameterDeclaration parameter)
        {
            _parameters.Add(parameter);
        }
    }
}
