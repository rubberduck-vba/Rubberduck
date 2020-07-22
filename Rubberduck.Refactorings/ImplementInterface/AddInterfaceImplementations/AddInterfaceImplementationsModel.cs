using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.AddInterfaceImplementations
{
    public class AddInterfaceImplementationsModel : IRefactoringModel
    {
        private readonly Dictionary<Declaration, string> _interfaceMemberContent;
        private readonly string _defaultImplementation;

        public AddInterfaceImplementationsModel(QualifiedModuleName targetModule, string interfaceName, IList<Declaration> members)
        {
            _interfaceMemberContent = new Dictionary<Declaration, string>();
            _defaultImplementation = $"    {Tokens.Err}.Raise 5 {Resources.Refactorings.Refactorings.ImplementInterface_TODO}";
            TargetModule = targetModule;
            InterfaceName = interfaceName;
            Members = members;
        }

        public QualifiedModuleName TargetModule { get; }
        public string InterfaceName { get; }
        public IList<Declaration> Members { get; }

        public void SetMemberImplementation(Declaration member, string implementation)
        {
            if (!Members.Contains(member))
            {
                throw new ArgumentException();
            }

            if (!_interfaceMemberContent.ContainsKey(member))
            {
                _interfaceMemberContent.Add(member, implementation);
                return;
            }
            _interfaceMemberContent[member] = implementation;
        }

        public string GetMemberImplementation(Declaration member)
        {
            if (!Members.Contains(member))
            {
                throw new ArgumentException();
            }

            if (_interfaceMemberContent.TryGetValue(member, out var implementation))
            {
                return string.IsNullOrEmpty(implementation) 
                    ? _defaultImplementation
                    : implementation;
            }

            return _defaultImplementation;
        }
    }
}