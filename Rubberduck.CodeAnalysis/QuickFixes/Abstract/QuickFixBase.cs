using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using NLog;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.CodeAnalysis.QuickFixes.Abstract
{
    internal abstract class QuickFixBase : IQuickFix
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

        /// <summary>
        /// FixMany defers the enumeration of inspection results to the QuickFix 
        /// </summary>
        /// <remarks>
        /// The default implementation enumerates the results collection calling Fix() for each result.
        /// Override this funcion when a QuickFix needs operate on results as a group (e.g., RemoveUnusedDeclarationQuickFix)
        /// </remarks>
        public virtual void Fix(IReadOnlyCollection<IInspectionResult> results, IRewriteSession rewriteSession)
        {
            foreach (var result in results)
            {
                Fix(result, rewriteSession);
            }
        }

        public abstract string Description(IInspectionResult result);

        public abstract bool CanFixMultiple { get; }
        public abstract bool CanFixInProcedure { get; }
        public abstract bool CanFixInModule { get; }
        public abstract bool CanFixInProject { get; }
        public abstract bool CanFixAll { get; }
    }
}