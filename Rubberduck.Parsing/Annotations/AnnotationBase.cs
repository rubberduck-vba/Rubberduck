using System;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class AnnotationBase : IAnnotation
    {
        public const string ANNOTATION_MARKER = "@";

        private readonly Lazy<int?> _annotatedLine;

        protected AnnotationBase(AnnotationType annotationType, QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context)
        {
            AnnotationType = annotationType;
            QualifiedSelection = qualifiedSelection;
            Context = context;
            _annotatedLine = new Lazy<int?>(GetAnnotatedLine);
        }

        public AnnotationType AnnotationType { get; }
        public QualifiedSelection QualifiedSelection { get; }
        public VBAParser.AnnotationContext Context { get; }

        public int? AnnotatedLine => _annotatedLine.Value;

        public virtual bool AllowMultiple { get; } = false;

        public override string ToString() => $"Annotation Type: {AnnotationType}";


        private int? GetAnnotatedLine()
        {
            var enclosingEndOfStatement = Context.GetAncestor<VBAParser.EndOfStatementContext>();

            //Annotations on the same line as non-whitespace statements do not scope to anything.
            if (enclosingEndOfStatement.Start.TokenIndex != 0)
            {
                var firstEndOfLine = enclosingEndOfStatement.GetFirstEndOfLine();
                var parentEndOfLine = Context.GetAncestor<VBAParser.EndOfLineContext>();
                if (firstEndOfLine.Equals(parentEndOfLine))
                {
                    return null;
                }
            }

            var lastToken = enclosingEndOfStatement.stop;
            return lastToken.Type == VBAParser.NEWLINE 
                   ? lastToken.Line + 1
                   : lastToken.Line;
        }
    }
}
