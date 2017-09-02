using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Abstract
{
    public class QuickFixBase
    {
        private HashSet<Type> _supportedInspections = new HashSet<Type>();
        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void RemoveInspections(params IInspection[] inspections)
        {
            _supportedInspections = _supportedInspections.Except(inspections.Select(s => s.Type)).ToHashSet();
        }

        public void RegisterInspections(params IInspection[] inspections)
        {
            _supportedInspections = inspections.Where(w => w != null).Select(s => s.Type).ToHashSet();
        }
    }
}