using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
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
            IEnumerable<IAnnotation> annotations,
            Accessibility accessibility, 
            DeclarationType declarationType,
            string value,
            ParserRuleContext context, 
            Selection selection, 
            bool isUserDefined = true)
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
                 isUserDefined,
                 annotations)
        {
            Expression = value;
        }

        public string Expression { get; }
    }
}
