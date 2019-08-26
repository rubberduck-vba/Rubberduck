using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Annotations
{
    public abstract class AnnotationBase : IAnnotation
    {
        public const string ANNOTATION_MARKER = "@";

        private readonly Lazy<int?> _annotatedLine;

        protected AnnotationBase(QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context)
        {
            QualifiedSelection = qualifiedSelection;
            Context = context;
            _annotatedLine = new Lazy<int?>(GetAnnotatedLine);
            MetaInformation = GetType().GetCustomAttributes(false).OfType<AnnotationAttribute>().Single();
        }
        
        public QualifiedSelection QualifiedSelection { get; }
        public VBAParser.AnnotationContext Context { get; }
        // sigh... we kinda want to seal this, but can't because it's not inherited from a class...
        public AnnotationAttribute MetaInformation { get; }
        public string AnnotationType => MetaInformation.Name;

        public int? AnnotatedLine => _annotatedLine.Value;

        public override string ToString() => $"Annotation Type: {GetType()}";

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
