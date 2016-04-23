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
        Control = 1 << 4,
        UserForm = 1 << 5,
        Document = 1 << 6,
        ModuleOption = 1 << 7,
        Member = 1 << 8,
        Procedure = 1 << 9 | Member,
        Function = 1 << 10 | Member,
        Property = 1 << 11 | Member,
        PropertyGet = 1 << 12 | Property | Function,
        PropertyLet = 1 << 13 | Property | Procedure,
        PropertySet = 1 << 14 | Property | Procedure,
        Parameter = 1 << 15,
        Variable = 1 << 16,
        Constant = 1 << 17,
        Enumeration = 1 << 18,
        EnumerationMember = 1 << 19 | Constant,
        Event = 1 << 20,
        UserDefinedType = 1 << 21,
        UserDefinedTypeMember = 1 << 22 | Variable,
        LibraryFunction = 1 << 23 | Function,
        LibraryProcedure = 1 << 24 | Procedure,
        LineLabel = 1 << 25
    }
}