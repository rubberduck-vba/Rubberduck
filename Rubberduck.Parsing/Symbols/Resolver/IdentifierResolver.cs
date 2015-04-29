using System.Linq;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols.Resolver
{
    public class IdentifierResolver : ResolverBase<ParserRuleContext>
    {
        public IdentifierResolver(Declarations declarations)
            : base(declarations)
        { }

        private static readonly Accessibility[] GlobalModifiers =
        {
            Accessibility.Global,
            Accessibility.Public,
            Accessibility.Implicit
        };

        public override Declaration Resolve(ParserRuleContext identifierContext, QualifiedModuleName currentModule, string currentScope, Declaration qualifier = null)
        {
            var identifier = identifierContext.GetText();
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