using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierResolver
    {
        private readonly Declarations _declarations;
        private readonly QualifiedModuleName _qualifiedParent;

        public IdentifierResolver(Declarations declarations, QualifiedModuleName qualifiedParent)
        {
            _declarations = declarations;
            _qualifiedParent = qualifiedParent;
        }

        public void Resolve(string identifierUsage, string currentScope, ParserRuleContext context, bool isAssignmentTarget = false, bool hasExplicitLetStatement = false)
        {
            var identifiers = identifierUsage.Split('.');
            var parentModuleName = _qualifiedParent;

            for (var i = 0; i < identifiers.Length; i++)
            {
                var identifier = identifiers[i];
                
                var isLeaf = (i == identifiers.Length - 1);
                
                var result = identifier == Tokens.Me 
                    ? Resolve(_declarations[_qualifiedParent.ComponentName], _qualifiedParent)
                    : Resolve(_declarations[identifier], currentScope, parentModuleName, isLeaf);

                parentModuleName = result != null 
                    ? result.QualifiedName.QualifiedModuleName 
                    : _qualifiedParent;

                if (result == null)
                {
                    break;
                }
                
                // bugger...
                //var reference = new IdentifierReference(parentModuleName, identifier, selection, context, result, isAssignmentTarget, hasExplicitLetStatement);
                //result.AddReference(reference);
            }
        }

        private static readonly DeclarationType[] LeafDeclarationTypes =
        {
            DeclarationType.Constant,
            DeclarationType.EnumerationMember,
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.LibraryFunction,
            DeclarationType.LibraryProcedure,
            DeclarationType.Parameter,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet,
            DeclarationType.UserDefinedTypeMember,
            DeclarationType.Variable
        };

        private Declaration Resolve(IEnumerable<Declaration> matches, QualifiedModuleName qualifiedParent)
        {
            // resolves "Me" token
            return matches.SingleOrDefault(match =>
                match.Project == qualifiedParent.Project && match.IdentifierName == qualifiedParent.ComponentName);
        }

        private Declaration Resolve(IEnumerable<Declaration> matches, string currentScope, QualifiedModuleName qualifiedParent, bool isLeaf)
        {
            var declarations = matches as List<Declaration> ?? matches.ToList();

            var currentScopeMatch = declarations.SingleOrDefault(match =>
                qualifiedParent.Project.Equals(match.Project) && match.Scope == currentScope
                && isLeaf || !LeafDeclarationTypes.Contains(match.DeclarationType));

            if (currentScopeMatch != null)
            {
                return currentScopeMatch;
            }

            var currentModuleMatch = declarations.SingleOrDefault(match =>
                qualifiedParent.Project.Equals(match.Project) && qualifiedParent.ComponentName == match.ComponentName
                && isLeaf || !LeafDeclarationTypes.Contains(match.DeclarationType));

            if (currentModuleMatch != null)
            {
                return currentModuleMatch;
            }

            var currentProjectMatch = declarations.SingleOrDefault(match =>
                qualifiedParent.Project.Equals(match.Project)
                && isLeaf || !LeafDeclarationTypes.Contains(match.DeclarationType));

            if (currentProjectMatch != null)
            {
                return currentProjectMatch;
            }

            var referencedProjects = qualifiedParent.Project.References.Cast<Reference>()
                .Where(reference => !reference.BuiltIn && !reference.IsBroken)
                .Select(reference =>
                    reference.VBE.VBProjects.Cast<VBProject>()
                    .SingleOrDefault(vbp => vbp.FileName == reference.FullPath))
                .ToList();

            var referencedProjectsMatch = declarations.SingleOrDefault(match =>
                referencedProjects.Contains(match.Project)
                && isLeaf || !LeafDeclarationTypes.Contains(match.DeclarationType));

            return referencedProjectsMatch;
        }
    }
}
