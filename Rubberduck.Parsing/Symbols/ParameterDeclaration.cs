using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class ParameterDeclaration : Declaration
    {
        private readonly bool _isOptional;
        private readonly bool _isByRef;
        private readonly bool _isArray;
        private readonly bool _isParamArray;

        /// <summary>
        /// Creates a new built-in parameter declaration.
        /// </summary>
        public ParameterDeclaration(QualifiedMemberName qualifiedName, 
            Declaration parentDeclaration, 
            string asTypeName,
            bool isOptional, 
            bool isByRef, 
            bool isArray = false, 
            bool isParamArray = false)
            : base(qualifiedName, parentDeclaration, parentDeclaration, asTypeName, false, false, Accessibility.Implicit, DeclarationType.Parameter, null, Selection.Home)
        {
            _isOptional = isOptional;
            _isByRef = isByRef;
            _isArray = isArray;
            _isParamArray = isParamArray;
        }

        /// <summary>
        /// Creates a new user declaration for a parameter.
        /// </summary>
        public ParameterDeclaration(QualifiedMemberName qualifiedName, 
            Declaration parentDeclaration,
            ParserRuleContext context, 
            Selection selection, 
            string asTypeName,
            bool isOptional,
            bool isByRef,
            bool isArray = false, 
            bool isParamArray = false)
            : base(qualifiedName, parentDeclaration, parentDeclaration, asTypeName, false, false, Accessibility.Implicit, DeclarationType.Parameter, context, selection, false)
        {
            _isOptional = isOptional;
            _isByRef = isByRef;
            _isArray = isArray;
            _isParamArray = isParamArray;
        }

        public bool IsOptional { get { return _isOptional; } }
        public bool IsByRef { get { return _isByRef; } }
        public override bool IsArray()
        {
            return _isArray;
        }

        public bool IsParamArray { get { return _isParamArray; } }
    }
}
