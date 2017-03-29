using Antlr4.Runtime;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class ParameterDeclaration : Declaration
    {
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
                  asTypeContext,
                  false)
        {
            IsOptional = isOptional;
            IsByRef = isByRef;
            IsImplicitByRef = false;
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
            bool isParamArray = false,
            bool isUserDefined = true)
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
                  isUserDefined)
        {
            IsOptional = isOptional;
            IsByRef = isByRef;
            IsImplicitByRef = isByRef && ((VBAParser.ArgContext) context)?.BYREF() == null;
            IsParamArray = isParamArray;
        }

        public ParameterDeclaration(ComParameter parameter, Declaration parent, QualifiedModuleName module)
            : this(
                module.QualifyMemberName(parameter.Name),
                parent,
                parameter.TypeName,
                null,
                null,
                parameter.IsOptional,
                parameter.IsByRef,
                parameter.IsArray,
                parameter.IsParamArray)
        { }
             
        public bool IsOptional { get; }
        public bool IsByRef { get; }
        public bool IsImplicitByRef { get; }
        public bool IsParamArray { get; set; }
    }
}
