using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Annotations
{
    public class ParseTreeAnnotation
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

        // FIXME annotation constructor for unit-testing purposes alone!
        internal ParseTreeAnnotation(IAnnotation annotation, QualifiedSelection qualifiedSelection, IEnumerable<string> arguments)
        {
            Annotation = annotation;
            QualifiedSelection = qualifiedSelection;
            _annotatedLine = new Lazy<int?>(() => null);
            Context = null;
            AnnotationArguments = arguments.ToList();
        }

        // needs to be accessible to IllegalAnnotationInspection
        public QualifiedSelection QualifiedSelection { get; }
        public VBAParser.AnnotationContext Context { get; }
        public int? AnnotatedLine => _annotatedLine.Value;

        // needs to be accessible to all external consumers
        public IAnnotation Annotation { get; }
        public IReadOnlyList<string> AnnotationArguments { get; }

        private static List<string> AnnotationParametersFromContext(VBAParser.AnnotationContext context)
        {
            var parameters = new List<string>();
            var argList = context?.annotationArgList();
            if (argList != null)
            {
                // CAUTION! THIS MUST NOT ADJUST THE QUOTING BEHAVIOUR!
                // the reason for that is the different quoting requirements for attributes.
                // some attributes require quoted values, some require unquoted values.
                // we currently don't have a mechanism to specify which needs which
                parameters.AddRange(argList.annotationArg().Select(arg => arg.GetText()));
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
