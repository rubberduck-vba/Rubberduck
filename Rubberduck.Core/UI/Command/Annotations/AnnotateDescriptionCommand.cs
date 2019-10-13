using System;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.Annotations
{
    public sealed class AnnotateDescriptionCommand : AnnotateCommandBase
    {
        public AnnotateDescriptionCommand(IRewritingManager manager, IAnnotationUpdater updater) 
            : base(manager, updater) { }

        protected override IAnnotation GetAnnotation(Declaration target)
            => new DescriptionAnnotation();

        protected override string[] GetAnnotationArgs()
        {
            throw new NotImplementedException();
        }
    }
}
