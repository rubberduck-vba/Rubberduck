using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ClassModuleDeclaration : Declaration
    {
        public ClassModuleDeclaration(
                  QualifiedMemberName qualifiedName,
                  Declaration projectDeclaration,
                  string name,
                  bool isBuiltIn,
                  IEnumerable<IAnnotation> annotations,
                  Attributes attributes)
            : base(
                  qualifiedName,
                  projectDeclaration,
                  projectDeclaration,
                  name,
                  false,
                  false,
                  Accessibility.Public,
                  DeclarationType.ClassModule,
                  null,
                  Selection.Home,
                  isBuiltIn,
                  annotations,
                  attributes)
        {
        }

        /// <summary>
        /// Gets an attribute value indicating whether a class is exposed to other projects.
        /// If this value is false, any public types and members cannot be accessed from outside the project they're declared in.
        /// </summary>
        public bool IsExposed
        {
            get
            {
                IEnumerable<string> value;
                if (Attributes.TryGetValue("VB_Exposed", out value))
                {
                    return value.Single() == "True";
                }
                return false;
            }
        }
    }
}
