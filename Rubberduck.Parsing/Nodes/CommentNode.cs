using System;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Nodes
{
    /// <summary>
    /// Represents a comment.
    /// </summary>
    /// <remarks>
    /// This is working around the limitations of the .g4 grammar file in use,
    /// which ignores comments. Ideally comments would be part of the grammar,
    /// and parsed along with the rest of the language syntax into an IParseTree.
    /// </remarks>
    public class CommentNode
    {
        private readonly string _comment;
        private readonly QualifiedSelection _qualifiedSelection;

        /// <summary>
        /// Creates a new comment node.
        /// </summary>
        /// <param name="comment">The comment line text, including the comment marker.</param>
        /// <param name="qualifiedSelection">The information required to locate and select this node in its VBE code pane.</param>
        public CommentNode(string comment, QualifiedSelection qualifiedSelection)
        {
            _comment = comment;
            _qualifiedSelection = qualifiedSelection;
        }

        /// <summary>
        /// Gets the comment line text, including the comment marker.
        /// </summary>
        public string Comment { get { return _comment; } }

        /// <summary>
        /// Gets the trimmed comment text, without the comment marker.
        /// </summary>
        public string CommentText { get { return _comment.Remove(_comment.IndexOf("'", StringComparison.Ordinal), 1).Trim(); } }

        /// <summary>
        /// The token used to indicate a comment.
        /// </summary>
        public string Marker
        {
            get
            {
                return _comment.StartsWith(Tokens.CommentMarker)
                                ? Tokens.CommentMarker
                                : Tokens.Rem;
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