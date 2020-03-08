using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

//This is the assembly name of the project Rubberduck.Main.
[assembly: InternalsVisibleTo("Rubberduck")]
[assembly: InternalsVisibleTo("RubberduckTests")]
// Moq needs this to access the types
[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2")]

[assembly: ComVisible(false)]
// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("26096055-4801-45f6-a82b-add86311eddf")]
