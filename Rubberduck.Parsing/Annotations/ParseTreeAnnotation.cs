using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Annotations
{
    public class ParseTreeAnnotation : IParseTreeAnnotation
    {
        public const string ANNOTATION_MARKER = "@";

        private readonly Lazy<int?> _annotatedLine;
        
        internal ParseTreeAnnotation(IAnnotation annotation, QualifiedSelection qualifiedSelection, VBAParser.AnnotationContext context)
        {
            QualifiedSelection = qualifiedSelection;
            Context = context;
            Annotation = annotation;
            _annotatedLine = new Lazy<int?>(GetAnnotatedLine);
            AnnotationArguments = AnnotationParametersFromContext(Context);
        }

        public QualifiedSelection QualifiedSelection { get; }
        public VBAParser.AnnotationContext Context { get; }
        public int? AnnotatedLine => _annotatedLine.Value;

        public IAnnotation Annotation { get; }
        public IReadOnlyList<string> AnnotationArguments { get; }

        private List<string> AnnotationParametersFromContext(VBAParser.AnnotationContext context)
        {
            var parameters = new List<string>();
            var argList = context?.annotationArgList();
            if (argList != null)
            {
                parameters.AddRange(Annotation.ProcessAnnotationArguments(argList.annotationArg().Select(arg => arg.GetText())));
            }
            return parameters;
        }

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
