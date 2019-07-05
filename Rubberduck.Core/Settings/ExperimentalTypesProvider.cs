using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Settings
{
    public interface IExperimentalTypesProvider
    {
        IReadOnlyList<Type> ExperimentalTypes { get; }
    }

    public class ExperimentalTypesProvider : IExperimentalTypesProvider
    {
        public IReadOnlyList<Type> ExperimentalTypes { get; }

        public ExperimentalTypesProvider(IEnumerable<Type> experimentalTypes)
        {
            ExperimentalTypes = experimentalTypes.ToList().AsReadOnly();
        }
    }
}
