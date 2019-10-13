using System;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.Annotations
{
    public sealed class AnnotateFolderCommand : AnnotateCommandBase
    {
        public AnnotateFolderCommand(IRewritingManager manager, IAnnotationUpdater updater) 
            : base(manager, updater) { }

        protected override string[] GetAnnotationArgs()
        {
            throw new NotImplementedException();
        }

        protected override IAnnotation GetAnnotation(Declaration target)
            => new FolderAnnotation();
    }
}
