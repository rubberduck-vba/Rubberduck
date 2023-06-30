using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

// ReSharper disable InconsistentNaming
// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.

namespace Rubberduck.UnitTesting
{
    internal class FileCopyParams : IFileCopyParams
    {
        public string Source { get; } = nameof(Source);

        public string Destination { get; } = nameof(Destination);
    }
}
