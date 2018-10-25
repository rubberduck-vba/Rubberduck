using System;
using System.Diagnostics;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public sealed class FallbackApp : IHostApplication
    {
        public FallbackApp(IVBE vbe)
        { }

        public string ApplicationName => "(unknown)";

        public void Dispose() { }
    }
}
