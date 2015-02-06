using System.Collections.Generic;
using Antlr4.Runtime.Tree;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;

namespace Rubberduck.VBA.Nodes
{
    public class VBComponentParseResult
    {
        public VBComponentParseResult(VBComponent component, IParseTree parseTree, IEnumerable<CommentNode> comments)
        {
            _component = component;
            _qualifiedName = new QualifiedModuleName(component.Collection.Parent.Name, component.Name);
            _parseTree = parseTree;
            _comments = comments;
        }

        private readonly VBComponent _component;
        public VBComponent Component { get { return _component; } }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private IParseTree _parseTree;
        public IParseTree ParseTree { get { return _parseTree; } }

        private IEnumerable<CommentNode> _comments;
        public IEnumerable<CommentNode> Comments { get { return _comments; } }
    }
}