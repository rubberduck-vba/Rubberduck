using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulatableField : IEncapsulateFieldRefactoringElement
    {
        string TargetID { get; }
        Declaration Declaration { get; }
        bool EncapsulateFlag { get; set; }
        string PropertyIdentifier { set; get; }
        string PropertyAsTypeName { set; get; }
        bool CanBeReadWrite { set; get; }
        bool ImplementLet { get; }
        bool ImplementSet { get; }
        bool IsReadOnly { set; get; }
        string ParameterName { get; }
        IValidateVBAIdentifiers NameValidator { set; get; }
        IEncapsulateFieldConflictFinder ConflictFinder { set; get; }
        bool TryValidateEncapsulationAttributes(out string errorMessage);
        string AccessorInProperty { get; }
        string AccessorLocalReference { get; }
        string AccessorExternalReference { /*set;*/ get; }
        IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets { get; }
    }

    public interface IUsingBackingField : IEncapsulatableField
    {
        string FieldIdentifier { set; get; }
        string FieldAsTypeName { set; get; }
        //string AccessorInProperty { get; }
        //string AccessorLocalReference { get; }
        //string AccessorExternalReference { set; get; }
    }

    public interface IConvertToUDTMember : IEncapsulatableField
    {
        //string AccessorInProperty { get; }
        //string AccessorLocalReference { get; }
        //string AccessorExternalReference { set; get; }
        string UDTMemberIdentifier { set; get; }
        string UDTMemberDeclaration { get; }
        IObjectStateUDT ObjectStateUDT { set; get; }
    }

    //public class EncapsulateField : IEncapsulatableField, IUsingBackingField//, IConvertToUDTMember
    //{
    //    private Declaration _declaration;
    //    public EncapsulateField(Declaration declaration)
    //    {
    //        _declaration = declaration;
    //    }

    //    public string IdentifierName => _declaration.IdentifierName;
    //    public string TargetID => _declaration.IdentifierName;
    //    public Declaration Declaration => _declaration;
    //    public QualifiedModuleName QualifiedModuleName => _declaration.QualifiedModuleName;
    //    public string AsTypeName => _declaration.AsTypeName;
    //    public bool EncapsulateFlag { set; get; }
    //    public string PropertyIdentifier { set; get; }
    //    public string PropertyAsTypeName { set; get; }
    //    public bool CanBeReadWrite { set; get; }
    //    public bool ImplementLet { get; }
    //    public bool ImplementSet { get; }
    //    public bool IsReadOnly { set; get; }
    //    public string ParameterName { get; }
    //    public string AccessorInProperty { set; get; }
    //    public string AccessorLocalReference { set; get; }
    //    public string AccessorExternalReference { set; get; }
    //    public string FieldIdentifier { set; get; }
    //    public string FieldAsTypeName { set; get; }
    //    public string UDTMemberIdentifier { set; get; }
    //    public string UDTMemberDeclaration { set; get; }
    //    public IObjectStateUDT Parent { set; get; }
    //}
}
