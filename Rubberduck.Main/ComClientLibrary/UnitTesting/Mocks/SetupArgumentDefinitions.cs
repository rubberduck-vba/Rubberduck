using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.ISetupArgumentDefinitionsGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface ISetupArgumentDefinitions : IEnumerable
    {
        [DispId(WellKnownDispIds.Value)]
        SetupArgumentDefinition Item(int Index);

        [DispId(1)]
        int Count { get; }

        [DispId(WellKnownDispIds.NewEnum)]
        IEnumerator _GetEnumerator();
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.SetupArgumentDefinitionsGuid),
        ProgId(RubberduckProgId.SetupArgumentDefinitionsProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(ISetupArgumentDefinitions))
    ]
    public class SetupArgumentDefinitions : ISetupArgumentDefinitions, IEnumerable<SetupArgumentDefinition>
    {
        private readonly List<SetupArgumentDefinition> _definitions;

        public SetupArgumentDefinitions()
        {
            _definitions = new List<SetupArgumentDefinition>();
        }

        public SetupArgumentDefinition Item(int Index) => _definitions.ElementAt(Index);

        public int Count => _definitions.Count;

        public IEnumerator<SetupArgumentDefinition> GetEnumerator() => _definitions.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => _definitions.GetEnumerator();
        
        public IEnumerator _GetEnumerator() => _definitions.GetEnumerator();

        internal void Add(SetupArgumentDefinition definition)
        {
            _definitions.Add(definition);
        }
    }
}
