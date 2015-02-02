using Rubberduck.Extensions;

namespace Rubberduck.VBA.Nodes
{
    public class CommentNode
    {
        public CommentNode(string comment, QualifiedSelection qualifiedSelection)
        {
            _comment = comment;
            _qualifiedSelection = qualifiedSelection;
        }

        private readonly string _comment;
        public string Comment { get { return _comment; } }

        public string Marker { get { return _comment[0] == '\'' ? "'" : ReservedKeywords.Rem; } }

        private readonly QualifiedSelection _qualifiedSelection;
        public QualifiedSelection QualifiedSelection { get { return _qualifiedSelection; } }
    }
}