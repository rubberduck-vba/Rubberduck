using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public class RubberduckParserState
    {
        // keys are the declarations; values indicate whether a declaration is resolved.
        private readonly ConcurrentDictionary<Declaration, ResolutionState> _declarations =
            new ConcurrentDictionary<Declaration, ResolutionState>();

        private readonly ConcurrentDictionary<VBComponent, ITokenStream> _tokenStreams =
            new ConcurrentDictionary<VBComponent, ITokenStream>();

        /// <summary>
        /// Gets all unresolved declarations.
        /// </summary>
        public IEnumerable<Declaration> UnresolvedDeclarations
        {
            get
            {
                return _declarations.Where(d => d.Value == ResolutionState.Unresolved)
                    .Select(d => d.Key);
            }
        }

        /// <summary>
        /// Gets a copy of the collected declarations of the specified <see cref="DeclarationType"/>.
        /// </summary>
        /// <param name="declarationType"></param>
        /// <returns></returns>
        public IEnumerable<Declaration> OfType(DeclarationType declarationType)
        {
            return AllDeclarations.Where(declaration =>
                declaration.DeclarationType == declarationType);
        }

        /// <summary>
        /// Gets a copy of the collected declarations of any one of the specified <see cref="DeclarationType"/> values.
        /// </summary>
        /// <param name="declarationTypes"></param>
        /// <returns></returns>
        public IEnumerable<Declaration> OfType(params DeclarationType[] declarationTypes)
        {
            return AllDeclarations.Where(declaration =>
                declarationTypes.Any(type => declaration.DeclarationType == type));
        }

        /// <summary>
        /// Gets a copy of the collected declarations containing all identifiers declared for the specified <see cref="component"/>.
        /// </summary>
        /// <param name="component"></param>
        /// <returns></returns>
        public IEnumerable<Declaration> Declarations(VBComponent component)
        {
            if (component == null)
            {
                throw new ArgumentNullException();
            }

            return AllDeclarations.Where(declaration =>
                declaration.QualifiedName.QualifiedModuleName.Component == component);
        }

        /// <summary>
        /// Gets a copy of the collected declarations containing all identifiers declared in or below the specified <see cref="scope"/>.
        /// </summary>
        /// <param name="scope"></param>
        /// <returns></returns>
        public IEnumerable<Declaration> Declarations(string scope = null)
        {
            var skip = string.IsNullOrEmpty(scope);
            return AllDeclarations.Where(declaration => skip || declaration.Scope.StartsWith(scope ?? string.Empty));
        }

        /// <summary>
        /// Gets a copy of the collected declarations.
        /// </summary>
        private IEnumerable<Declaration> AllDeclarations { get { return _declarations.Keys.ToList(); } }

        /// <summary>
        /// Adds the specified <see cref="Declaration"/> to the collection.
        /// </summary>
        /// <param name="declaration"></param>
        /// <returns>Returns true when successful, replaces existing key reference.</returns>
        public bool AddDeclaration(Declaration declaration)
        {
            if (!_declarations.TryAdd(declaration, ResolutionState.Unresolved))
            {
                if (RemoveDeclaration(declaration))
                {
                    return _declarations.TryAdd(declaration, ResolutionState.Unresolved);
                }
            }

            return false;
        }

        public bool AddTokenStream(VBComponent component, ITokenStream stream)
        {
            return _tokenStreams.TryAdd(component, stream);
        }

        /// <summary>
        /// Removes the specified <see cref="declaration"/> from the collection.
        /// </summary>
        /// <param name="declaration"></param>
        /// <returns>Returns true when successful.</returns>
        private bool RemoveDeclaration(Declaration declaration)
        {
            foreach (var reference in declaration.References)
            {
                MarkForResolution(reference.ParentScope);
            }
            foreach (var reference in declaration.MemberCalls)
            {
                MarkForResolution(reference.ParentScope);
            }

            ResolutionState state;
            return _declarations.TryRemove(declaration, out state);
        }

        public void MarkForResolution(string scope)
        {
            foreach (var declaration in _declarations.Keys.Where(d => !d.IsDirty && (d.Scope == scope || d.ParentScope == scope)))
            {
                declaration.IsDirty = true;
            }
        }
    }
}