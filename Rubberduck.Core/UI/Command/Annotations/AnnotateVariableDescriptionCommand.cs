using System;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.Annotations
{
    public sealed class AnnotateVariableDescriptionCommand : AnnotateCommandBase
    {
        public AnnotateVariableDescriptionCommand(IRewritingManager manager, IAnnotationUpdater updater) 
            : base(manager, updater) { }

        protected override IAnnotation GetAnnotation(Declaration target)
            => new VariableDescriptionAnnotation();

        protected override string[] GetAnnotationArgs()
        {
            throw new NotImplementedException();
        }
    }
}
