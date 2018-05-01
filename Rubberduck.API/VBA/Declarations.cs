using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.SharedResources.COM;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IDeclarationsGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IDeclarations : IEnumerable
    {
        [DispId(0)]
        Declaration Item(int Index);

        [DispId(1)]
        int Count { get; }

        [DispId(-4)]
        IEnumerator _GetEnumerator();
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.DeclarationsClassGuid),
        ProgId(RubberduckProgId.DeclarationsProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IDeclarations))
    ]
    public class Declarations : IDeclarations, IEnumerable<Declaration>
    {
        private readonly IEnumerable<Declaration> _declarations;

        public int Count => _declarations.Count();

        private Declarations(IEnumerable<Declaration> declarations)
        {
            _declarations = declarations;
        }

        public Declaration Item(int Index)
        {
            return _declarations.ElementAt(Index);
        }

        public IEnumerator<Declaration> GetEnumerator()
        {
            return _declarations.GetEnumerator();
        }

        public IEnumerator _GetEnumerator()
        {
            return _declarations.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _declarations.GetEnumerator();
        }
    }
}
