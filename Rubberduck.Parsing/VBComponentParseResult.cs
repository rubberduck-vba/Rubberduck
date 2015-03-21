using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Nodes;

namespace Rubberduck.Parsing
{
    public class VBComponentParseResult
    {
        public VBComponentParseResult(VBComponent component, IParseTree parseTree, IEnumerable<CommentNode> comments, ParserRuleContext context = null)
        {
            _component = component;
            _qualifiedName = component.QualifiedName();
            _parseTree = parseTree;
            _comments = comments;
            _context = context;
        }

        private readonly VBComponent _component;
        public VBComponent Component { get { return _component; } }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private IParseTree _parseTree;
        public IParseTree ParseTree { get { return _parseTree; } }

        private IEnumerable<CommentNode> _comments;
        public IEnumerable<CommentNode> Comments { get { return _comments; } }

        private ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }
    }
}