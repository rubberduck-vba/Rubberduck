using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ExternalProcedureDeclaration : Declaration, IDeclarationWithParameter
    {
        private readonly List<Declaration> _parameters;

        public ExternalProcedureDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            DeclarationType declarationType,
            string asTypeName,
            Accessibility accessibility,
            ParserRuleContext context,
            Selection selection,
            bool isBuiltIn,
            IEnumerable<IAnnotation> annotations)
            : base(
                  name,
                  parent,
                  parentScope,
                  asTypeName,
                  false,
                  false,
                  accessibility,
                  declarationType,
                  context,
                  selection,
                  isBuiltIn,
                  annotations,
                  null)
        {
            _parameters = new List<Declaration>();
        }

        public IEnumerable<Declaration> Parameters
        {
            get
            {
                return _parameters.ToList();
            }
        }

        public void AddParameter(Declaration parameter)
        {
            _parameters.Add(parameter);
        }
    }
}
