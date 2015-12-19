using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.Parsing.VBA
{
    public class RubberduckParserState
    {
        // keys are the declarations; values indicate whether a declaration is resolved.
        private readonly ConcurrentHashSet<Declaration> _declarations =
            new ConcurrentHashSet<Declaration>();

        private readonly ConcurrentDictionary<VBComponent, ITokenStream> _tokenStreams =
            new ConcurrentDictionary<VBComponent, ITokenStream>();

        private readonly ConcurrentDictionary<VBComponent, IParseTree> _parseTrees =
            new ConcurrentDictionary<VBComponent, IParseTree>();

        public event EventHandler StateChanged;

        private void OnStateChanged()
        {
            var handler = StateChanged;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }

        private readonly ConcurrentDictionary<VBComponent, ParserState> _moduleStates =
            new ConcurrentDictionary<VBComponent, ParserState>();

        private readonly ConcurrentDictionary<VBComponent, SyntaxErrorException> _moduleExceptions =
            new ConcurrentDictionary<VBComponent, SyntaxErrorException>();

        public void SetModuleState(VBComponent component, ParserState state, SyntaxErrorException parserError = null)
        {
            _moduleStates[component] = state;
            _moduleExceptions[component] = parserError;

            Status = _moduleStates.Values.Any(value => value == ParserState.Error)
                ? ParserState.Error
                : _moduleStates.Values.Any(value => value == ParserState.Parsing)
                    ? ParserState.Parsing
                    : _moduleStates.Values.Any(value => value == ParserState.Resolving)
                        ? ParserState.Resolving
                        : ParserState.Ready;

        }

        public ParserState GetModuleState(VBComponent component)
        {
            return _moduleStates[component];
        }

        private ParserState _status;
        public ParserState Status { get { return _status; } private set { if(_status != value) {_status = value; OnStateChanged();} } }

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

        private readonly ConcurrentDictionary<VBComponent, IEnumerable<CommentNode>> _comments =
            new ConcurrentDictionary<VBComponent, IEnumerable<CommentNode>>();

        public IEnumerable<CommentNode> Comments
        {
            get 
            {
                return _comments.Values.SelectMany(comments => comments);
            }
        }

        public void SetModuleComments(VBComponent component, IEnumerable<CommentNode> comments)
        {
            _comments[component] = comments;
        }

        /// <summary>
        /// Gets a copy of the collected declarations.
        /// </summary>
        public IEnumerable<Declaration> AllDeclarations { get { return _declarations; } }

        /// <summary>
        /// Adds the specified <see cref="Declaration"/> to the collection (replaces existing).
        /// </summary>
        public void AddDeclaration(Declaration declaration)
        {
            if (_declarations.Add(declaration))
            {
                return;
            }

            if (RemoveDeclaration(declaration))
            {
                _declarations.Add(declaration);
            }
        }

        public void ClearDeclarations(VBComponent component)
        {
            var declarations = _declarations.Where(k =>
                k.QualifiedName.QualifiedModuleName.Project == component.Collection.Parent
                && k.ComponentName == component.Name);

            while (true)
            {
                try
                {
                    foreach (var declaration in declarations)
                    {
                        _declarations.Remove(declaration);
                    }

                    return;
                }
                catch (InvalidOperationException)
                {
                    
                }
            }
        }

        public void AddTokenStream(VBComponent component, ITokenStream stream)
        {
            _tokenStreams[component] = stream;
        }

        public void AddParseTree(VBComponent component, IParseTree parseTree)
        {
            _parseTrees[component] = parseTree;
        }

        public IEnumerable<KeyValuePair<VBComponent, IParseTree>> ParseTrees { get { return _parseTrees; } }

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
            return _declarations.Remove(declaration);
        }

        /// <summary>
        /// Ensures parser state accounts for built-in declarations.
        /// This method has no effect if built-in declarations have already been loaded.
        /// </summary>
        public void AddBuiltInDeclarations(IHostApplication hostApplication)
        {
            if (_declarations.Any(declaration => declaration.IsBuiltIn))
            {
                return;
            }

            var builtInDeclarations = VbaStandardLib.Declarations;

            // cannot be strongly-typed here because of constraints on COM interop and generics in the inheritance hierarchy. </rant>
            if (hostApplication /*is ExcelApp*/ .ApplicationName == "Excel") 
            {
                builtInDeclarations = builtInDeclarations.Concat(ExcelObjectModel.Declarations);
            }

            foreach (var declaration in builtInDeclarations)
            {
                AddDeclaration(declaration);
            }
        }
    }
}