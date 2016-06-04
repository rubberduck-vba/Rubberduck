using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class ParameterDeclaration : Declaration
    {
        private readonly bool _isOptional;
        private readonly bool _isByRef;

        /// <summary>
        /// Creates a new built-in parameter declaration.
        /// </summary>
        public ParameterDeclaration(QualifiedMemberName qualifiedName, 
            Declaration parentDeclaration, 
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            bool isOptional, 
            bool isByRef, 
            bool isArray = false, 
            bool isParamArray = false)
            : base(
                  qualifiedName, 
                  parentDeclaration, 
                  parentDeclaration, 
                  asTypeName,
                  typeHint,
                  false, 
                  false, 
                  Accessibility.Implicit,
                  DeclarationType.Parameter, 
                  null, 
                  Selection.Home,
                  isArray,
                  asTypeContext)
        {
            _isOptional = isOptional;
            _isByRef = isByRef;
            IsParamArray = isParamArray;
        }

        /// <summary>
        /// Creates a new user declaration for a parameter.
        /// </summary>
        public ParameterDeclaration(QualifiedMemberName qualifiedName, 
            Declaration parentDeclaration,
            ParserRuleContext context, 
            Selection selection, 
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            bool isOptional,
            bool isByRef,
            bool isArray = false, 
            bool isParamArray = false)
            : base(
                  qualifiedName, 
                  parentDeclaration, 
                  parentDeclaration,
                  asTypeName,
                  typeHint,
                  false, 
                  false, 
                  Accessibility.Implicit,
                  DeclarationType.Parameter, 
                  context, 
                  selection,
                  isArray,
                  asTypeContext,
                  false)
        {
            _isOptional = isOptional;
            _isByRef = isByRef;
            IsParamArray = isParamArray;
        }

        public bool IsOptional { get { return _isOptional; } }
        public bool IsByRef { get { return _isByRef; } }
        public bool IsParamArray { get; set; }
    }
}
