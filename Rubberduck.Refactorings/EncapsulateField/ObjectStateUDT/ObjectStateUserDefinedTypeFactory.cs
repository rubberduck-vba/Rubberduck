using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor;
using System;

namespace Rubberduck.Refactorings
{
    public interface IObjectStateUserDefinedTypeFactory
    {
        IObjectStateUDT Create(QualifiedModuleName qualifiedModuleName);
        IObjectStateUDT Create(IUserDefinedTypeCandidate userDefinedTypeField);
    }

    public class ObjectStateUserDefinedTypeFactory : IObjectStateUserDefinedTypeFactory
    {
        public IObjectStateUDT Create(QualifiedModuleName qualifiedModuleName)
        {
            return new ObjectStateUDT(qualifiedModuleName);
        }

        public IObjectStateUDT Create(IUserDefinedTypeCandidate userDefinedTypeField)
        {
            if ((userDefinedTypeField.Declaration.AsTypeDeclaration?.Accessibility ?? Accessibility.Implicit) != Accessibility.Private)
            {
                throw new ArgumentException();
            }

            return new ObjectStateUDT(userDefinedTypeField);
        }
    }
}
