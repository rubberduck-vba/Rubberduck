using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.AddInterfaceImplementations
{
    public class AddInterfaceImplementationsModel : IRefactoringModel
    {
        public QualifiedModuleName TargetModule { get; }
        public string InterfaceName { get; }
        public IList<Declaration> Members { get; }

        public AddInterfaceImplementationsModel(QualifiedModuleName targetModule, string interfaceName, IList<Declaration> members)
        {
            TargetModule = targetModule;
            InterfaceName = interfaceName;
            Members = members;
        }
    }
}