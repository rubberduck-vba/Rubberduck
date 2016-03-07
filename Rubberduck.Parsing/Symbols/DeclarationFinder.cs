using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class DeclarationFinder
    {
        private readonly IDictionary<QualifiedModuleName, CommentNode[]> _comments;
        private readonly IDictionary<string, Declaration[]> _declarationsByName;

        public DeclarationFinder(IReadOnlyList<Declaration> declarations, IEnumerable<CommentNode> comments)
        {
            _comments = comments.GroupBy(node => node.QualifiedSelection.QualifiedName)
                .ToDictionary(grouping => grouping.Key, grouping => grouping.ToArray());

            _declarationsByName = declarations.GroupBy(declaration => declaration.IdentifierName)
                .ToDictionary(grouping => grouping.Key, grouping => grouping.ToArray());
        }

        private readonly HashSet<Accessibility> _projectScopePublicModifiers =
            new HashSet<Accessibility>(new[]
            {
                Accessibility.Public,
                Accessibility.Global,
                Accessibility.Friend,
                Accessibility.Implicit,
            });

        public IEnumerable<CommentNode> ModuleComments(QualifiedModuleName module)
        {
            CommentNode[] result;
            if (_comments.TryGetValue(module, out result))
            {
                return result;
            }

            return new List<CommentNode>();
        }

        public IEnumerable<Declaration> MatchTypeName(string name)
        {
            return MatchName(name).Where(declaration =>
                declaration.DeclarationType == DeclarationType.Class ||
                declaration.DeclarationType == DeclarationType.UserDefinedType);
        }

        public IEnumerable<Declaration> MatchName(string name)
        {
            Declaration[] result;
            if (_declarationsByName.TryGetValue(name, out result))
            {
                return result;
            }

            return new List<Declaration>();
        }

        public Declaration FindProject(Declaration currentScope, string name)
        {
            Declaration result = null;
            try
            {
                result = MatchName(name).SingleOrDefault(project => project.DeclarationType == DeclarationType.Project 
                    && (currentScope == null || project.Project == currentScope.Project));
            }
            catch (InvalidOperationException exception)
            {
                Debug.WriteLine("Multiple matches found for project '{0}'.\n{1}", name, exception);
            }

            return result;
        }

        public Declaration FindStdModule(Declaration parent, string name, bool includeBuiltIn = false)
        {
            Declaration result = null;
            try
            {
                result = MatchName(name).SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Module
                    && (parent == null || parent.Equals(declaration.ParentDeclaration))
                    && (includeBuiltIn || !declaration.IsBuiltIn));
            }
            catch (InvalidOperationException exception)
            {
                Debug.WriteLine("Multiple matches found for std.module '{0}'.\n{1}", name, exception);
            }

            return result;
        }

        public Declaration FindUserDefinedType(Declaration parent, string name, bool includeBuiltIn = false)
        {
            Declaration result = null;
            try
            {
                result = MatchName(name).SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.UserDefinedType
                    && parent == null
                        ? _projectScopePublicModifiers.Contains(declaration.Accessibility)
                        : parent.Equals(declaration.ParentDeclaration)
                          && (includeBuiltIn || !declaration.IsBuiltIn));
            }
            catch (InvalidOperationException exception)
            {
                Debug.WriteLine("Multiple matches found for user-defined type '{0}'.\n{1}", name, exception);
            }

            return result;
        }

        public Declaration FindClass(Declaration parent, string name, bool includeBuiltIn = false)
        {
            if (parent == null)
            {
                throw new ArgumentNullException("parent");
            }

            Declaration result = null;
            try
            {
                result = MatchName(name).SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Class
                    && parent.Equals(declaration.ParentDeclaration)
                    && (includeBuiltIn || !declaration.IsBuiltIn));
            }
            catch (InvalidOperationException exception)
            {
                Debug.WriteLine("Multiple matches found for class '{0}'.\n{1}", name, exception);
            }

            return result;
        }
    }
}