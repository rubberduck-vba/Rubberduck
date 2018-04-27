using System.Runtime.InteropServices;
using Source = Rubberduck.Parsing.Symbols;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true), 
        Guid(RubberduckGuid.DeclarationTypeGuid)
    ]
    public enum DeclarationType
    {
        BracketedExpression = Source.DeclarationType.BracketedExpression,
        ClassModule = Source.DeclarationType.ClassModule,
        ComAlias = Source.DeclarationType.ComAlias,
        Constant = Source.DeclarationType.Constant,
        Control = Source.DeclarationType.Control,
        Document = Source.DeclarationType.Document,
        Enumeration = Source.DeclarationType.Enumeration,
        EnumerationMember = Source.DeclarationType.EnumerationMember,
        Event = Source.DeclarationType.Event,
        Function = Source.DeclarationType.Function,
        LibraryFunction = Source.DeclarationType.LibraryFunction,
        LibraryProcedure = Source.DeclarationType.LibraryProcedure,
        LineLabel = Source.DeclarationType.LineLabel,
        Member = Source.DeclarationType.Member,
        Module = Source.DeclarationType.Module,
        Parameter = Source.DeclarationType.Parameter,
        ProceduralModule = Source.DeclarationType.ProceduralModule,
        Procedure = Source.DeclarationType.Procedure,
        Project = Source.DeclarationType.Project,
        Property = Source.DeclarationType.Property,
        PropertyGet = Source.DeclarationType.PropertyGet,
        PropertyLet = Source.DeclarationType.PropertyLet,
        PropertySet = Source.DeclarationType.PropertySet,
        UnresolvedMember = Source.DeclarationType.UnresolvedMember,
        UserDefinedType = Source.DeclarationType.UserDefinedType,
        UserDefinedTypeMember = Source.DeclarationType.UserDefinedTypeMember,
        UserForm = Source.DeclarationType.UserForm,
        Variable = Source.DeclarationType.Variable
    }
}
