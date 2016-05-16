using System;

namespace Rubberduck.Parsing.Symbols
{
    [Flags]
    public enum DeclarationType
    {
        Project = 1 << 0,
        Module = 1 << 1,
        ProceduralModule = 1 << 2 | Module,
        ClassModule = 1 << 3 | Module,
        UserForm = 1 << 4,
        Document = 1 << 5,
        ModuleOption = 1 << 6,
        Member = 1 << 7,
        Procedure = 1 << 8 | Member,
        Function = 1 << 9 | Member,
        Property = 1 << 10 | Member,
        PropertyGet = 1 << 11 | Property | Function,
        PropertyLet = 1 << 12 | Property | Procedure,
        PropertySet = 1 << 13 | Property | Procedure,
        Parameter = 1 << 14,
        Variable = 1 << 15,
        Control = 1 << 16 | Variable,
        Constant = 1 << 17,
        Enumeration = 1 << 18,
        EnumerationMember = 1 << 19,
        Event = 1 << 20,
        UserDefinedType = 1 << 21,
        UserDefinedTypeMember = 1 << 22,
        LibraryFunction = 1 << 23 | Function,
        LibraryProcedure = 1 << 24 | Procedure,
        LineLabel = 1 << 25
    }
}
