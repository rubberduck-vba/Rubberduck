using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Symbols.Resolver
{
    public class AmbiguousIdentifierResolver : ResolverBase<VBAParser.AmbiguousIdentifierContext>
    {
        public AmbiguousIdentifierResolver(Declarations declarations)
            :base(declarations) { }

        private static readonly Accessibility[] GlobalModifiers =
        {
            Accessibility.Global,
            Accessibility.Public,
            Accessibility.Implicit
        };

        public override Declaration Resolve(VBAParser.AmbiguousIdentifierContext context, QualifiedModuleName currentModule, string currentScope, Declaration qualifier = null)
        {
            var identifier = context.GetText();
            var matches = Declarations.Items.Where(item => item.IdentifierName == identifier).ToList();

            if (matches.Any(item => item.Scope == currentScope))
            {
                // multiple matches in current scope would be ambiguous.
                return matches.Single(item => item.Scope == currentScope);
            }

            if (matches.Any(item => currentModule.Project.Equals(item.Project) && item.ComponentName == currentModule.ComponentName))
            {
                // multiple matches in current module would be ambiguous.
                return matches.Single(item => currentModule.Project.Equals(item.Project) && item.ComponentName == currentModule.ComponentName);
            }

            if (matches.Any(item => currentModule.Project.Equals(item.Project) 
                && (GlobalModifiers.Contains(item.Accessibility) || item.Accessibility == Accessibility.Friend)))
            {
                if (qualifier == null)
                {
                    // unqualified multiple matches in current project would be ambiguous.
                    return matches.Single(item => currentModule.Project.Equals(item.Project) 
                        && (GlobalModifiers.Contains(item.Accessibility) || item.Accessibility == Accessibility.Friend));
                }
                
                // qualified multiple matches would be ambiguous.
                return matches.Single(item => currentModule.Project.Equals(item.Project) 
                    && (GlobalModifiers.Contains(item.Accessibility) || item.Accessibility == Accessibility.Friend)
                    && item.ParentScope == qualifier.Scope);
            }

            if (qualifier == null)
            {
                // multiple matches in global scope would be ambiguous.   
                return matches.SingleOrDefault(item => GlobalModifiers.Contains(item.Accessibility));
            }

            // qualified multiple matches in global scope would be ambiguous.
            return matches.SingleOrDefault(item => GlobalModifiers.Contains(item.Accessibility)
                && item.ParentScope == qualifier.Scope);
        }
    }
}