using System.Runtime.InteropServices;

namespace Rubberduck.API
{
    [ComVisible(true)]
    //[Flags]
    public enum DeclarationType
    {
        Project, //= 1 << 0,
        StandardModule, //= 1 << 1,
        ClassModule,// = 1 << 2,
        Control, //= 1 << 3,
        UserForm,// = 1 << 4,
        Document,// = 1 << 5,
        ModuleOption,// = 1 << 6,
        Procedure, //= 1 << 8,
        Function,// = 1 << 9,
        PropertyGet,// = 1 << 11,
        PropertyLet, //= 1 << 12,
        PropertySet, //= 1 << 13,
        Parameter, //= 1 << 14,
        Variable, //= 1 << 15,
        Constant,// = 1 << 16,
        Enumeration, //= 1 << 17,
        EnumerationMember, //= 1 << 18,
        Event, //= 1 << 19,
        UserDefinedType,// = 1 << 20,
        UserDefinedTypeMember,// = 1 << 21,
        LibraryFunction,// = 1 << 22,
        LibraryProcedure,// = 1 << 23,
        LineLabel,// = 1 << 24,
        //Member = Procedure | Function | PropertyGet | PropertyLet | PropertySet,
        //Property = PropertyGet | PropertyLet | PropertySet,
        //Module = StandardModule | ClassModule | UserForm | Document
    }
}
