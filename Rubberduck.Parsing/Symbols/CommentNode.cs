using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    /// <summary>
    /// Represents a comment.
    /// </summary>
    public class CommentNode
    {
        private readonly string _comment;
        private readonly string _marker;
        private readonly QualifiedSelection _qualifiedSelection;

        /// <summary>
        /// Creates a new comment node.
        /// </summary>
        /// <param name="comment">The comment line text, without the comment marker.</param>
        /// <param name="qualifiedSelection">The information required to locate and select this node in its VBE code pane.</param>
        public CommentNode(string comment, string marker, QualifiedSelection qualifiedSelection)
        {
            _comment = comment;
            _marker = marker;
            _qualifiedSelection = qualifiedSelection;
        }

        /// <summary>
        /// Gets the comment text, without the comment marker.
        /// </summary>
        public string CommentText { get { return _comment; } }

        /// <summary>
        /// The token used to indicate a comment.
        /// </summary>
        public string Marker
        {
            get
            {
                return _marker;
            }
        }

        /// <summary>
        /// Gets the information required to locate and select this node in its VBE code pane.
        /// </summary>
        public QualifiedSelection QualifiedSelection { get { return _qualifiedSelection; } }

        /// <summary>
        /// Returns the comment text.
        /// </summary>
        public override string ToString()
        {
            return _comment;
        }
    }
}
