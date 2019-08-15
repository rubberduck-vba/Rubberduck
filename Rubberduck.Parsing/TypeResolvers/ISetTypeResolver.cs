using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.TypeResolvers
{
    public interface ISetTypeResolver
    {
        /// <summary>
        /// Determines the declaration representing the Set type of an expression, if there is one.
        /// </summary>
        /// <returns>
        /// Declaration representing the Set type of an expression, if there is such a declaration, and
        /// null, otherwise. In particular, null is returned for expressions of Set type Variant and Object.  
        /// </returns>
        Declaration SetTypeDeclaration(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule);

        /// <summary>
        /// Determines the name of the Set type of an expression, if it has a Set type.
        /// </summary>
        /// <returns>
        /// Qualified name of the Set type of the expression, if there is one, and
        /// NotAnObject, otherwise. Returns null, if the resolution fails.
        /// </returns>
        string SetTypeName(VBAParser.ExpressionContext expression, QualifiedModuleName containingModule);
    }
}
