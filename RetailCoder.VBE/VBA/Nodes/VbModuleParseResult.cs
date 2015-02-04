using System.Collections.Generic;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections;

namespace Rubberduck.VBA.Nodes
{
    public class VbModuleParseResult
    {
        public VbModuleParseResult(QualifiedModuleName qualifiedName, IParseTree parseTree, IEnumerable<CommentNode> comments)
        {
            _qualifiedName = qualifiedName;
            _parseTree = parseTree;
            _comments = comments;
        }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private IParseTree _parseTree;
        public IParseTree ParseTree { get { return _parseTree; } }

        private IEnumerable<CommentNode> _comments;
        public IEnumerable<CommentNode> Comments { get { return _comments; } }

    }
}