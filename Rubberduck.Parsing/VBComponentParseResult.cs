using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    public class VBComponentParseResult
    {
        public VBComponentParseResult(VBComponent component, IParseTree parseTree, IEnumerable<CommentNode> comments, ITokenStream tokenStream)
        {
            _component = component;
            _qualifiedName = new QualifiedModuleName(component);
            _parseTree = parseTree;
            _comments = comments;
            _tokenStream = tokenStream;

            var listener = new DeclarationSymbolsListener(_qualifiedName, Accessibility.Implicit, _component.Type);
            var walker = new ParseTreeWalker();
            walker.Walk(listener, _parseTree);

            _declarations.AddRange(listener.Declarations.Items);
        }

        private readonly List<Declaration> _declarations = new List<Declaration>();
        public IEnumerable<Declaration> Declarations { get { return _declarations; } }

        private readonly VBComponent _component;
        public VBComponent Component { get { return _component; } }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private IParseTree _parseTree;
        public IParseTree ParseTree { get { return _parseTree; } }

        private IEnumerable<CommentNode> _comments;
        public IEnumerable<CommentNode> Comments { get { return _comments; } }

        private readonly ITokenStream _tokenStream;
        public TokenStreamRewriter GetRewriter()
        {
            return new TokenStreamRewriter(_tokenStream);
        }
    }
}