using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.ComReflection.TypeLibReflection
{
    public interface IRegisteredLibraryFinderService
    {
        IEnumerable<RegisteredLibraryInfo> FindRegisteredLibraries();
        bool TryGetRegisteredLibraryInfo(Guid typeLibGuid, out RegisteredLibraryInfo info);
    }
}