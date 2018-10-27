using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IQuickFix
    {
        void Fix(IInspectionResult result, IRewriteSession rewriteSession = null);
        string Description(IInspectionResult result);

        bool CanFixInProcedure { get; }
        bool CanFixInModule { get; }
        bool CanFixInProject { get; }

        IReadOnlyCollection<Type> SupportedInspections { get; }

        void RegisterInspections(params Type[] inspections);
        void RemoveInspections(params Type[] inspections);
    }
}