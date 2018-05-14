using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;
using Source = Rubberduck.Parsing.Symbols;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true), 
        Guid(RubberduckGuid.DeclarationTypeGuid)
    ]
    public enum DeclarationType
    {
        //TODO: figure out a cleaner way to handle long/int 
        BracketedExpression = (int)Source.DeclarationType.BracketedExpression,
        ClassModule = (int)Source.DeclarationType.ClassModule,
        ComAlias = (int)Source.DeclarationType.ComAlias,
        Constant = (int)Source.DeclarationType.Constant,
        Control = (int)Source.DeclarationType.Control,
        Document = (int)Source.DeclarationType.Document,
        Enumeration = (int)Source.DeclarationType.Enumeration,
        EnumerationMember = (int)Source.DeclarationType.EnumerationMember,
        Event = (int)Source.DeclarationType.Event,
        Function = (int)Source.DeclarationType.Function,
        LibraryFunction = (int)Source.DeclarationType.LibraryFunction,
        LibraryProcedure = (int)Source.DeclarationType.LibraryProcedure,
        LineLabel = (int)Source.DeclarationType.LineLabel,
        Member = (int)Source.DeclarationType.Member,
        Module = (int)Source.DeclarationType.Module,
        Parameter = (int)Source.DeclarationType.Parameter,
        ProceduralModule = (int)Source.DeclarationType.ProceduralModule,
        Procedure = (int)Source.DeclarationType.Procedure,
        Project = (int)Source.DeclarationType.Project,
        Property = (int)Source.DeclarationType.Property,
        PropertyGet = (int)Source.DeclarationType.PropertyGet,
        PropertyLet = (int)Source.DeclarationType.PropertyLet,
        PropertySet = (int)Source.DeclarationType.PropertySet,
        UnresolvedMember = (int)Source.DeclarationType.UnresolvedMember,
        UserDefinedType = (int)Source.DeclarationType.UserDefinedType,
        UserDefinedTypeMember = (int)Source.DeclarationType.UserDefinedTypeMember,
        UserForm = (int)Source.DeclarationType.UserForm,
        Variable = (int)Source.DeclarationType.Variable
    }
}
