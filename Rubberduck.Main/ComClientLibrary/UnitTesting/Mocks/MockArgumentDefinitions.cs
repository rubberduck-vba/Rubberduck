using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IMockArgumentDefinitionsGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IMockArgumentDefinitions : IEnumerable
    {
        [DispId(WellKnownDispIds.Value)]
        MockArgumentDefinition Item(int Index);

        [DispId(1)]
        int Count { get; }

        [DispId(WellKnownDispIds.NewEnum)]
        IEnumerator _GetEnumerator();
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.MockArgumentDefinitionsGuid),
        ProgId(RubberduckProgId.MockArgumentDefinitionsProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IMockArgumentDefinitions))
    ]
    public class MockArgumentDefinitions : IMockArgumentDefinitions, IEnumerable<MockArgumentDefinition>
    {
        private readonly List<MockArgumentDefinition> _definitions;

        public MockArgumentDefinitions()
        {
            _definitions = new List<MockArgumentDefinition>();
        }

        public MockArgumentDefinition Item(int Index) => _definitions.ElementAt(Index);

        public int Count => _definitions.Count;

        public IEnumerator<MockArgumentDefinition> GetEnumerator() => _definitions.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => _definitions.GetEnumerator();
        
        public IEnumerator _GetEnumerator() => _definitions.GetEnumerator();

        internal void Add(MockArgumentDefinition definition)
        {
            _definitions.Add(definition);
        }
    }
}
