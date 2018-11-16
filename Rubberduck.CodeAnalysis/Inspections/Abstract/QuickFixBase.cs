using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class QuickFixBase : IQuickFix
    {
        protected readonly ILogger Logger = LogManager.GetCurrentClassLogger();
        private HashSet<Type> _supportedInspections;
        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        protected QuickFixBase(params Type[] inspections)
        {
            RegisterInspections(inspections);
        }

        public void RegisterInspections(params Type[] inspections)
        {
            if (!inspections.All(s => s.GetInterfaces().Any(a => a == typeof(IInspection))))
            {
#if DEBUG
                throw new ArgumentException($"Parameters must implement {nameof(IInspection)}", nameof(inspections));
#else
                inspections.Where(s => s.GetInterfaces().All(i => i != typeof(IInspection))).ToList()
                    .ForEach(i => Logger.Error($"Type {i.Name} does not implement {nameof(IInspection)}"));
#endif
            }

            _supportedInspections = inspections.ToHashSet();
        }

        public void RemoveInspections(params Type[] inspections)
        {
            _supportedInspections = _supportedInspections.Except(inspections).ToHashSet();
        }

        public virtual CodeKind TargetCodeKind => CodeKind.CodePaneCode;

        public abstract void Fix(IInspectionResult result, IRewriteSession rewriteSession);
        public abstract string Description(IInspectionResult result);

        public abstract bool CanFixInProcedure { get; }
        public abstract bool CanFixInModule { get; }
        public abstract bool CanFixInProject { get; }
    }
}