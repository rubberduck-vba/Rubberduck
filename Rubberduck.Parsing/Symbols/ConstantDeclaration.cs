using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class ConstantDeclaration : Declaration
    {
        public ConstantDeclaration(
            QualifiedMemberName qualifiedName, 
            Declaration parentDeclaration, 
            string parentScope,
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            Accessibility accessibility, 
            DeclarationType declarationType, 
            string value)
            : this(
                  qualifiedName,
                  parentDeclaration, 
                  parentScope, 
                  asTypeName,
                  asTypeContext,
                  typeHint,
                  accessibility, 
                  declarationType, 
                  value, 
                  null, 
                  Selection.Home,
                  true)
        {
        }

        public ConstantDeclaration(
            QualifiedMemberName qualifiedName, 
            Declaration parentDeclaration,
            string parentScope,
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            Accessibility accessibility, 
            DeclarationType declarationType,
            string value,
            ParserRuleContext context, 
            Selection selection, 
            bool isBuiltIn = false)
            :base(
                 qualifiedName, 
                 parentDeclaration, 
                 parentScope, 
                 asTypeName, 
                 typeHint,
                 true, 
                 false, 
                 accessibility,
                 declarationType, 
                 context, 
                 selection,
                 false,
                 asTypeContext,
                 isBuiltIn)
        {
            _expression = value;
        }

        private readonly string _expression;
        public string Expression { get { return _expression; } }
    }
}
