using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

//This is the assembly name of the project Rubberduck.Main.
[assembly: InternalsVisibleTo("Rubberduck")]
[assembly: InternalsVisibleTo("RubberduckTests")]

//Allow Rubberduck.VBEditor.* projects to use internal class
[assembly: InternalsVisibleTo("Rubberduck.VBEditor.VB6")]
[assembly: InternalsVisibleTo("Rubberduck.VBEditor.VBA")]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("4758424b-266e-4a3a-83c6-4aa50af7ea9e")]
[assembly: ComVisible(false)]