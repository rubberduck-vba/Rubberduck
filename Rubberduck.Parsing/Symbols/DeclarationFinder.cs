using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public class DeclarationFinder
    {
        private readonly IDictionary<QualifiedModuleName, CommentNode[]> _comments;
        private readonly IDictionary<QualifiedModuleName, IAnnotation[]> _annotations;
        private readonly IReadOnlyList<Declaration> _declarations;
        private readonly IDictionary<string, Declaration[]> _declarationsByName;

        public DeclarationFinder(
            IReadOnlyList<Declaration> declarations,
            IEnumerable<CommentNode> comments,
            IEnumerable<IAnnotation> annotations)
        {
            _comments = comments.GroupBy(node => node.QualifiedSelection.QualifiedName)
                .ToDictionary(grouping => grouping.Key, grouping => grouping.ToArray());
            _annotations = annotations.GroupBy(node => node.QualifiedSelection.QualifiedName)
                .ToDictionary(grouping => grouping.Key, grouping => grouping.ToArray());
            _declarations = declarations;
            _declarationsByName = declarations.GroupBy(declaration => new
            {
                IdentifierName = declaration.Project != null &&
                        declaration.DeclarationType.HasFlag(DeclarationType.Project)
                            ? declaration.Project.Name.ToLowerInvariant()
                            : declaration.IdentifierName.ToLowerInvariant()
            })
            .ToDictionary(grouping => grouping.Key.IdentifierName, grouping => grouping.ToArray());
        }

        private readonly HashSet<Accessibility> _projectScopePublicModifiers =
            new HashSet<Accessibility>(new[]
            {
                Accessibility.Public,
                Accessibility.Global,
                Accessibility.Friend,
                Accessibility.Implicit,
            });

        public IEnumerable<Declaration> FindDeclarationsWithNonBaseAsType()
        {
            return _declarations
                .Where(d => 
                !string.IsNullOrWhiteSpace(d.AsTypeName) 
                && !d.AsTypeIsBaseType
                && d.DeclarationType != DeclarationType.Project
                && d.DeclarationType != DeclarationType.ProceduralModule).ToList();
        }

        public IEnumerable<CommentNode> ModuleComments(QualifiedModuleName module)
        {
            CommentNode[] result;
            if (_comments.TryGetValue(module, out result))
            {
                return result;
            }

            return new List<CommentNode>();
        }

        public IEnumerable<IAnnotation> ModuleAnnotations(QualifiedModuleName module)
        {
            IAnnotation[] result;
            if (_annotations.TryGetValue(module, out result))
            {
                return result;
            }

            return new List<IAnnotation>();
        }

        public IEnumerable<Declaration> MatchTypeName(string name)
        {
            return MatchName(name).Where(declaration =>
                declaration.DeclarationType.HasFlag(DeclarationType.ClassModule) ||
                declaration.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ||
                declaration.DeclarationType.HasFlag(DeclarationType.Enumeration));
        }

        public bool IsMatch(string declarationName, string potentialMatchName)
        {
            return string.Equals(declarationName, potentialMatchName, StringComparison.OrdinalIgnoreCase);
        }

        public Declaration FindEvent(Declaration module, string eventName)
        {
            var matches = MatchName(eventName);
            return matches.Where(m => module.Equals(Declaration.GetModuleParent(m)) && m.DeclarationType == DeclarationType.Event).FirstOrDefault();
        }

        public Declaration FindLabel(Declaration procedure, string label)
        {
            var matches = MatchName(label);
            return matches.Where(m => procedure.Equals(m.ParentDeclaration) && m.DeclarationType == DeclarationType.LineLabel).FirstOrDefault();
        }

        public IEnumerable<Declaration> MatchName(string name)
        {
            var normalizedName = name.ToLowerInvariant();
            Declaration[] result;
            if (_declarationsByName.TryGetValue(normalizedName, out result))
            {
                return result;
            }
            if (_declarationsByName.TryGetValue("_" + normalizedName, out result))
            {
                return result;
            }
            if (_declarationsByName.TryGetValue("i" + normalizedName, out result))
            {
                return result;
            }
            if (_declarationsByName.TryGetValue("_i" + normalizedName, out result))
            {
                return result;
            }

            return new List<Declaration>();
        }

        public Declaration FindProject(string name, Declaration currentScope = null)
        {
            Declaration result = null;
            try
            {
                result = MatchName(name).SingleOrDefault(project => project.DeclarationType.HasFlag(DeclarationType.Project)
                    && (currentScope == null || project.ProjectId == currentScope.ProjectId));
            }
            catch (InvalidOperationException exception)
            {
                Debug.WriteLine("Multiple matches found for project '{0}'.\n{1}", name, exception);
            }

            return result;
        }

        public Declaration FindStdModule(string name, Declaration parent = null, bool includeBuiltIn = false)
        {
            Declaration result = null;
            try
            {
                var matches = MatchName(name);
                result = matches.SingleOrDefault(declaration => declaration.DeclarationType.HasFlag(DeclarationType.ProceduralModule)
                    && (parent == null || parent.Equals(declaration.ParentDeclaration))
                    && (includeBuiltIn || !declaration.IsBuiltIn));
            }
            catch (InvalidOperationException exception)
            {
                Debug.WriteLine("Multiple matches found for std.module '{0}'.\n{1}", name, exception);
            }

            return result;
        }

        public Declaration FindUserDefinedType(string name, Declaration parent = null, bool includeBuiltIn = false)
        {
            Declaration result = null;
            try
            {
                var matches = MatchName(name);
                result = matches.SingleOrDefault(declaration => declaration.DeclarationType.HasFlag(DeclarationType.UserDefinedType)
                    && (parent == null || parent.Equals(declaration.ParentDeclaration))
                    && (includeBuiltIn || !declaration.IsBuiltIn));
            }
            catch (Exception exception)
            {
                Debug.WriteLine("Multiple matches found for user-defined type '{0}'.\n{1}", name, exception);
            }

            return result;
        }

        public Declaration FindEnum(string name, Declaration parent = null, bool includeBuiltIn = false)
        {
            Declaration result = null;
            try
            {
                var matches = MatchName(name);
                result = matches.SingleOrDefault(declaration => declaration.DeclarationType.HasFlag(DeclarationType.Enumeration)
                    && (parent == null || parent.Equals(declaration.ParentDeclaration))
                    && (includeBuiltIn || !declaration.IsBuiltIn));
            }
            catch (Exception exception)
            {
                Debug.WriteLine("Multiple matches found for enum type '{0}'.\n{1}", name, exception);
            }

            return result;
        }

        public Declaration FindClass(Declaration parent, string name, bool includeBuiltIn = false)
        {
            Declaration result = null;
            try
            {
                result = MatchName(name).SingleOrDefault(declaration => declaration.DeclarationType.HasFlag(DeclarationType.ClassModule)
                    && parent == null || parent.Equals(declaration.ParentDeclaration)
                    && (includeBuiltIn || !declaration.IsBuiltIn));
            }
            catch (InvalidOperationException exception)
            {
                Debug.WriteLine("Multiple matches found for class '{0}'.\n{1}", name, exception);
            }

            return result;
        }

        public Declaration FindReferencedProject(Declaration callingProject, string referencedProjectName)
        {
            return FindInReferencedProjectByPriority(callingProject, referencedProjectName, p => p.DeclarationType.HasFlag(DeclarationType.Project));
        }

        public Declaration FindModuleEnclosingProjectWithoutEnclosingModule(Declaration callingProject, Declaration callingModule, string calleeModuleName, DeclarationType moduleType)
        {
            var nameMatches = MatchName(calleeModuleName);
            var moduleMatches = nameMatches.Where(m => 
                m.DeclarationType.HasFlag(moduleType)
                && Declaration.GetProjectParent(m).Equals(callingProject)
                && !m.Equals(callingModule));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindModuleReferencedProject(Declaration callingProject, Declaration callingModule, string calleeModuleName, DeclarationType moduleType)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName, p => p.DeclarationType.HasFlag(moduleType));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindModuleReferencedProject(Declaration callingProject, Declaration callingModule, Declaration referencedProject, string calleeModuleName, DeclarationType moduleType)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName, p => referencedProject.Equals(Declaration.GetProjectParent(p)) && p.DeclarationType.HasFlag(moduleType));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberWithParent(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m =>
                m.DeclarationType.HasFlag(memberType)
                && callingParent.Equals(m.ParentDeclaration));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberEnclosingModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m =>
                m.DeclarationType.HasFlag(memberType)
                && Declaration.GetProjectParent(m).Equals(callingProject)
                && callingModule.Equals(Declaration.GetModuleParent(m)));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberEnclosingProcedure(Declaration enclosingProcedure, string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m =>
                m.DeclarationType.HasFlag(memberType)
                && enclosingProcedure.Equals(m.ParentDeclaration));
            var match = memberMatches.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberEnclosedProjectWithoutEnclosingModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            return FindMemberEnclosedProjectWithoutEnclosingModule(callingProject, callingModule, callingParent, memberName, memberType, DeclarationType.Module);
        }

        public Declaration FindMemberEnclosedProjectWithoutEnclosingModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType, DeclarationType moduleType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m =>
                m.DeclarationType.HasFlag(memberType)
                // TODO: Fix this?
                // Assume no module = built-in, not checkable thus we conservatively include it in the matches.
                && (Declaration.GetModuleParent(m) == null || Declaration.GetModuleParent(m).DeclarationType.HasFlag(moduleType))
                && Declaration.GetProjectParent(m).Equals(callingProject)
                && !callingModule.Equals(Declaration.GetModuleParent(m)));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberEnclosedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration memberModule, string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m =>
                m.DeclarationType.HasFlag(memberType)
                && Declaration.GetProjectParent(m).Equals(callingProject)
                && memberModule.Equals(Declaration.GetModuleParent(m)));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProject(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, DeclarationType moduleType, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType) && (Declaration.GetModuleParent(p) == null || Declaration.GetModuleParent(p).DeclarationType == moduleType));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInGlobalClassModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType) && (Declaration.GetModuleParent(p) == null || Declaration.GetModuleParent(p).DeclarationType == DeclarationType.ClassModule) && ((ClassModuleDeclaration)Declaration.GetModuleParent(p)).IsGlobalClassModule);
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration memberModule, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType) && memberModule.Equals(Declaration.GetModuleParent(p)));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProject(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration referencedProject, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType) && referencedProject.Equals(Declaration.GetProjectParent(p)));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        private Declaration FindInReferencedProjectByPriority(Declaration enclosingProject, string name, Func<Declaration, bool> predicate)
        {
            return FindAllInReferencedProjectByPriority(enclosingProject, name, predicate).FirstOrDefault();
        }

        private IEnumerable<Declaration> FindAllInReferencedProjectByPriority(Declaration enclosingProject, string name, Func<Declaration, bool> predicate)
        {
            var interprojectMatches = MatchName(name).Where(predicate).ToList();
            var projectReferences = ((ProjectDeclaration)enclosingProject).ProjectReferences.ToList();
            if (interprojectMatches.Count == 0)
            {
                yield break;
            }
            foreach (var projectReference in projectReferences)
            {
                var match = interprojectMatches.FirstOrDefault(interprojectMatch => interprojectMatch.ProjectId == projectReference.ReferencedProjectId);
                if (match != null)
                {
                    yield return match;
                }
            }
            yield break;
        }
    }
}