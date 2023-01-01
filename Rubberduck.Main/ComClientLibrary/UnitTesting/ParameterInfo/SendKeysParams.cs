using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

// ReSharper disable InconsistentNaming
// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.

namespace Rubberduck.UnitTesting
{
    internal class SendKeysParams : ISendKeysParams
    {
        public string Keys { get; } = nameof(Keys);

        public string Wait { get; } = nameof(Wait);
    }
}
