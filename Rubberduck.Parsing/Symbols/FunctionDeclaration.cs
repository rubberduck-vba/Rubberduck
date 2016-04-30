using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class FunctionDeclaration : Declaration, IDeclarationWithParameter, ICanBeDefaultMember
    {
        private readonly List<Declaration> _parameters;

        public FunctionDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            Accessibility accessibility,
            ParserRuleContext context,
            Selection selection,
            bool isBuiltIn,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(
                  name,
                  parent,
                  parentScope,
                  asTypeName,
                  false,
                  false,
                  accessibility,
                  DeclarationType.Function,
                  context,
                  selection,
                  isBuiltIn,
                  annotations,
                  attributes)
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

        public void Add(Declaration parameter)
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
