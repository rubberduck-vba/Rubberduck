using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

//This is the assembly name of the project Rubberduck.Main.
[assembly: InternalsVisibleTo("Rubberduck")]
// internals visible for testing and mocking
[assembly: InternalsVisibleTo("RubberduckTests")]
[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("ac4f1d22-d74b-45ff-ab0c-cc2a104fe023")]