using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// A special version of <see cref="TypeInfoVariablesCollection"/> which stores
    /// only the constants for the type library.
    /// </summary>
    /// <remarks>
    /// This aims to help work around the VBE type infos non-complaint
    /// behaviors of storing variables as a variable in a standard module
    /// which is considered a <see cref="TYPEKIND.TKIND_MODULE"/>. According
    /// to MS-OAUT open specifications "3.7.1.2 TYPEKIND Dependent Automation
    /// Type Description Elements", and "2.2.19 VARKIND Variable Kind Constants",
    /// the module should only contain a constant or a static variable which points
    /// to a coclass. In both cases, it is assumed that vardesc->lpvarValue->vt will
    /// always be valid. But that is not the case with a VBA type info representing a
    /// VBA standard module. Thus, this class is used to filter out all non constants
    /// from a module and thus conforms to the MS-OAUT specifications.
    ///
    /// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/7b1b8bd1-a067-4edb-9d72-6aa500d035a3
    /// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/a0e9d463-51a2-49cc-8935-a65c9338d3df
    /// </remarks>
    internal class TypeInfoConstantsCollection : TypeInfoVariablesCollection
    {
        private Dictionary<int, int> _mapper;

        public TypeInfoConstantsCollection(ITypeInfo parent, TYPEATTR attributes) : 
            base(parent, attributes)
        {
            // External consumers won't actually know the real index of the underlying
            // VARDESC struct and will expect to be able to enumerate with a for loop.
            // We map the original index to a new position to make it easy to enumerate
            // the constants regardless of their positions in the actual index. 
            _mapper = new Dictionary<int, int>();

            for (var i = 0; i < attributes.cVars; i++)
            {
                parent.GetVarDesc(i, out var ppVarDesc);
                var varDesc = StructHelper.ReadStructureUnsafe<VARDESC>(ppVarDesc);

                // VBA constants are "static".... go figure.
                if (varDesc.IsValidVBAConstant())
                {
                    _mapper.Add(_mapper.Count, i);
                }
                parent.ReleaseVarDesc(ppVarDesc);
            }
        }

        public override int Count => _mapper.Count;

        public override ITypeInfoVariable GetItemByIndex(int index) => 
            new TypeInfoVariable(Parent, _mapper[index]);

        public int MappedIndex(int index) => _mapper[index];
    }
}
