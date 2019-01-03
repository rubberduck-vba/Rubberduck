using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using Rubberduck.Resources.Registration;

[assembly: InternalsVisibleTo("RubberduckTests")]
[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2")]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid(RubberduckGuid.RubberduckTypeLibGuid)]
[assembly: ComVisible(false)]