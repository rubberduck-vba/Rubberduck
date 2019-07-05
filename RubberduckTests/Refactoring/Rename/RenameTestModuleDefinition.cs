using System;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;

namespace RubberduckTests.Refactoring.Rename
{
    internal struct RenameTestModuleDefinition
    {
        private CodeString _codeString;
        private string _expected;

        public string InputWithFauxCursor { get; private set; }

        public string Input
        {
            set
            {
                InputWithFauxCursor = value.Contains(RenameTests.FAUX_CURSOR) ? value : InputWithFauxCursor;
                _codeString = value.ToCodeString();
            }
            get => _codeString.Code;
        }

        public string Expected
        {
            set => _expected = value;
            get => _expected.Equals(string.Empty) ? Input : _expected;
        }

        public string ModuleName;
        public ComponentType ModuleType;
        public bool CheckExpectedEqualsActual;
        public List<string> ControlNames;
        public string NewName;
        public Selection? RenameSelection => _codeString.CaretPosition.ToOneBased();

        public RenameTestModuleDefinition(string moduleName, ComponentType moduleType = ComponentType.ClassModule)
        {
            _codeString = new CodeString(string.Empty, default);
            InputWithFauxCursor = string.Empty;
            _expected = string.Empty;
            ModuleName = moduleName;
            NewName = moduleName;
            ModuleType = moduleType;
            CheckExpectedEqualsActual = true;
            ControlNames = new List<String>();
        }
    }
}
