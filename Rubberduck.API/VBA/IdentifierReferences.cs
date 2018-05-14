using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IIdentifierReferencesGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IIdentifierReferences : IEnumerable
    {
        [DispId(WellKnownDispIds.Value)]
        IdentifierReference Item(int Index);

        [DispId(1)]
        int Count { get; }

        [DispId(WellKnownDispIds.NewEnum)]
        IEnumerator _GetEnumerator();
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.IdentifierReferencesClassGuid),
        ProgId(RubberduckProgId.IdentifierReferencesProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IIdentifierReferences))
    ]
    public class IdentifierReferences : IIdentifierReferences, IEnumerable<IdentifierReference>
    {
        private readonly IEnumerable<IdentifierReference> _references;

        public int Count => _references.Count();

        internal IdentifierReferences(IEnumerable<IdentifierReference> references)
        {
            _references = references;
        }

        public IdentifierReference Item(int Index)
        {
            return _references.ElementAt(Index);
        }

        public IEnumerator<IdentifierReference> GetEnumerator()
        {
            return _references.GetEnumerator();
        }

        public IEnumerator _GetEnumerator()
        {
            return _references.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _references.GetEnumerator();
        }
    }
}
