using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.VBA
{
    public class RubberduckParserState
    {
        public enum State
        {
            /// <summary>
            /// Parser state is in sync with the actual code in the VBE.
            /// </summary>
            Ready,
            /// <summary>
            /// One or more modules were modified, but parsing hasn't started yet.
            /// </summary>
            Dirty,
            /// <summary>
            /// Code from modified modules is being parsed.
            /// </summary>
            Parsing,
            /// <summary>
            /// Resolving identifier references.
            /// </summary>
            Resolving,
        }

        // keys are the declarations; values indicate whether a declaration is resolved.
        private readonly ConcurrentDictionary<Declaration, ResolutionState> _declarations =
            new ConcurrentDictionary<Declaration, ResolutionState>();

        private readonly ConcurrentDictionary<VBComponent, ITokenStream> _tokenStreams =
            new ConcurrentDictionary<VBComponent, ITokenStream>();

        public State Status { get; internal set; }

        /// <summary>
        /// Gets all unresolved declarations.
        /// </summary>
        public IEnumerable<Declaration> UnresolvedDeclarations
        {
            get
            {
                return _declarations.Select(d => d.Key);
            }
        }

        /// <summary>
        /// Gets a copy of the collected declarations containing all identifiers declared for the specified <see cref="component"/>.
        /// </summary>
        public IEnumerable<Declaration> Declarations(VBComponent component)
        {
            if (component == null)
            {
                throw new ArgumentNullException();
            }

            return AllDeclarations.Where(declaration =>
                declaration.QualifiedName.QualifiedModuleName.Component == component);
        }

        private IEnumerable<QualifiedContext> _obsoleteCallContexts = new List<QualifiedContext>();

        /// <summary>
        /// Gets <see cref="ParserRuleContext"/> objects representing 'Call' statements in the parse tree.
        /// </summary>
        public IEnumerable<QualifiedContext> ObsoleteCallContexts
        {
            get { return _obsoleteCallContexts; }
            internal set { _obsoleteCallContexts = value; }
        }

        private IEnumerable<QualifiedContext> _obsoleteLetContexts = new List<QualifiedContext>();

        /// <summary>
        /// Gets <see cref="ParserRuleContext"/> objects representing explicit 'Let' statements in the parse tree.
        /// </summary>
        public IEnumerable<QualifiedContext> ObsoleteLetContexts
        {
            get { return _obsoleteLetContexts; }
            internal set { _obsoleteLetContexts = value; }
        }

        private IEnumerable<CommentNode> _comments = new List<CommentNode>(); 
        public IEnumerable<CommentNode> Comments { get { return _comments; } internal set { _comments = value; } }

        /// <summary>
        /// Gets a copy of the collected declarations.
        /// </summary>
        public IEnumerable<Declaration> AllDeclarations { get { return _declarations.Keys.ToList(); } }

        /// <summary>
        /// Adds the specified <see cref="Declaration"/> to the collection (replaces existing).
        /// </summary>
        public void AddDeclaration(Declaration declaration)
        {
            if (_declarations.TryAdd(declaration, ResolutionState.Unresolved))
            {
                return;
            }

            if (RemoveDeclaration(declaration))
            {
                _declarations.TryAdd(declaration, ResolutionState.Unresolved);
            }
        }

        public void ClearDeclarations(VBComponent component)
        {
            var declarations = _declarations.Keys.Where(k =>
                k.QualifiedName.QualifiedModuleName.Project == component.Collection.Parent
                && k.ComponentName == component.Name);

            foreach (var declaration in declarations)
            {
                ResolutionState state;
                _declarations.TryRemove(declaration, out state);
            }
        }

        public void AddTokenStream(VBComponent component, ITokenStream stream)
        {
            _tokenStreams.TryAdd(component, stream);
        }

        public TokenStreamRewriter GetRewriter(VBComponent component)
        {
            return new TokenStreamRewriter(_tokenStreams[component]);
        }

        /// <summary>
        /// Removes the specified <see cref="declaration"/> from the collection.
        /// </summary>
        /// <param name="declaration"></param>
        /// <returns>Returns true when successful.</returns>
        private bool RemoveDeclaration(Declaration declaration)
        {
            ResolutionState state;
            return _declarations.TryRemove(declaration, out state);
        }
    }
}