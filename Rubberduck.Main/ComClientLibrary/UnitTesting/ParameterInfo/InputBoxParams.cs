using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

// ReSharper disable InconsistentNaming
// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.

namespace Rubberduck.UnitTesting
{
    internal class InputBoxParams : IInputBoxParams
    {
        public string Prompt { get; } = nameof(Prompt);
        public string Title { get; } = nameof(Title);
        public string Default { get; } = nameof(Default);
        public string XPos { get; } = nameof(XPos);
        public string YPos { get; } = nameof(YPos);
        public string HelpFile { get; } = nameof(HelpFile);
        public string Context { get; } = nameof(Context);
    }
}
