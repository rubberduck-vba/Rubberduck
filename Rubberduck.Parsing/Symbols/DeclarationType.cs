using System;

namespace Rubberduck.Parsing.Symbols
{
    [Flags]
    public enum DeclarationType
    {
        Project = 1 << 0,
        Module = 1 << 1,
        Class = 1 << 2,
        Control = 1 << 3,
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
        Constant = 1 << 16,
        Enumeration = 1 << 17,
        EnumerationMember = 1 << 18 | Constant,
        Event = 1 << 19,
        UserDefinedType = 1 << 20,
        UserDefinedTypeMember = 1 << 21 | Variable,
        LibraryFunction = 1 << 22 | Function,
        LibraryProcedure = 1 << 23 | Procedure,
        LineLabel = 1 << 24
    }
}