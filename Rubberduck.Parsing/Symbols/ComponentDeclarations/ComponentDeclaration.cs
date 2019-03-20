using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public abstract class ComponentDeclaration : Declaration
    {
        protected ComponentDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
            DeclarationType declarationType,
            bool isUserDefined,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes,
            bool isWithEvents = false)
            : base(
                qualifiedName,
                projectDeclaration,
                projectDeclaration,
                name,
                null,
                false,
                isWithEvents,
                Accessibility.Public,
                declarationType,
                null,
                null,
                Selection.Home,
                false,
                null,
                isUserDefined,
                annotations,
                attributes)
        {
        }        
    }
}
