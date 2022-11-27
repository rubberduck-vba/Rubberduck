using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

// ReSharper disable InconsistentNaming
// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.

namespace Rubberduck.UnitTesting
{
    internal class MsgBoxParams : IMsgBoxParams
    {
        public string Prompt { get; } = nameof(Prompt);
        public string Buttons { get; } = nameof(Buttons);
        public string Title { get; } = nameof(Title);
        public string HelpFile { get; } = nameof(HelpFile);
        public string Context { get; } = nameof(Context);
    }
}
