using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IQuickFix
    {
        void Fix(IInspectionResult result);
        string Description(IInspectionResult result);

        bool CanFixInProcedure { get; }
        bool CanFixInModule { get; }
        bool CanFixInProject { get; }

        IReadOnlyCollection<Type> SupportedInspections { get; }

        void RegisterInspections(params Type[] inspections);
        void RemoveInspections(params Type[] inspections);
    }
}