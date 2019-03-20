using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public class UserFormDeclaration : ClassModuleDeclaration
    {
        public UserFormDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(qualifiedName,
                projectDeclaration,
                name,
                DeclarationType.UserForm,
                true,
                annotations,
                attributes,
                true,
                true,
                false)
        { }
    }
}
