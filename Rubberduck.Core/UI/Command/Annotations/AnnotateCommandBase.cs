using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.Annotations
{
    public abstract class AnnotateCommandBase : CommandBase
    {
        private readonly IRewritingManager _manager;
        private readonly IAnnotationUpdater _updater;

        protected AnnotateCommandBase(IRewritingManager manager, IAnnotationUpdater updater) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _manager = manager;
            _updater = updater;
        }

        protected virtual IAnnotation GetAnnotation(Declaration target)
        {
            /* no-op */
            return default;
        }

        protected virtual IAnnotation GetAnnotation(IdentifierReference target)
        {
            /* no-op */
            return default;
        }

        protected virtual string[] GetAnnotationArgs() => Array.Empty<string>();

        private bool Annotate(Declaration target, IAnnotationUpdater updater, IRewriteSession session)
        {
            var annotation = GetAnnotation(target);
            if (annotation == default)
            {
                return false;
            }

            var args = GetAnnotationArgs();

            var existing = target.Annotations.FirstOrDefault(a => a.GetType() == annotation.GetType());
            if (existing == default)
            {
                updater.AddAnnotation(session, target, annotation, args);
            }
            else
            {
                updater.UpdateAnnotation(session, existing, annotation, args);
            }

            return true;
        }

        private bool Annotate(IdentifierReference target, IAnnotationUpdater updater, IRewriteSession session)
        {
            var annotation = GetAnnotation(target);
            if (annotation == default)
            {
                return false;
            }

            var args = GetAnnotationArgs();

            var existing = target.Annotations.FirstOrDefault(a => a.GetType() == annotation.GetType());
            if (existing == null)
            {
                updater.AddAnnotation(session, target, annotation, args);
            }
            else
            {
                updater.UpdateAnnotation(session, existing, annotation, args);
            }

            return true;
        }

        protected sealed override void OnExecute(object parameter)
        {
            if (!(parameter is Declaration target))
            {
                return;
            }

            var session = _manager.CheckOutCodePaneSession();
            Annotate(target, _updater, session);
            if (!session.TryRewrite())
            {
                Logger.Info("TryRewrite failed, annotation may not have been added.");
            }
        }
    }
}
