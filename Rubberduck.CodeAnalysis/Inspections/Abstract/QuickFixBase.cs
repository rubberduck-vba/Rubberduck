using System;
using System.Collections.Generic;
using System.Diagnostics;
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
                var dieNow = false;
                MustThrowException(ref dieNow);
                if (dieNow)
                {
                    throw new ArgumentException($"Parameters must implement {nameof(IInspection)}",
                        nameof(inspections));
                }

                inspections.Where(s => s.GetInterfaces().All(i => i != typeof(IInspection))).ToList()
                    .ForEach(i => Logger.Error($"Type {i.Name} does not implement {nameof(IInspection)}"));
            }

            _supportedInspections = inspections.ToHashSet();
        }

        // ReSharper disable once RedundantAssignment : conditional must be void but we can use ref
        [Conditional("DEBUG")] 
        private static void MustThrowException(ref bool dieNow)
        {
            dieNow = true;
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