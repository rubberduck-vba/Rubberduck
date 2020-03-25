using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

//This is the assembly name of the project Rubberduck.Main.
[assembly: InternalsVisibleTo("Rubberduck")]
// internals visible for testing and mocking
[assembly: InternalsVisibleTo("RubberduckTests")]
[assembly: InternalsVisibleTo("DynamicProxyGenAssembly2")]

[assembly: ComVisible(false)]
// Die folgende GUID bestimmt die ID der Typbibliothek, wenn dieses Projekt für COM verfügbar gemacht wird
[assembly: Guid("7f136926-696e-4051-bd40-efc19c8f78c6")]