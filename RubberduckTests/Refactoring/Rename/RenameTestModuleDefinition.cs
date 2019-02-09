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
        private string _inputWithFauxCursor;
        private string _expected;

        public string Input_WithFauxCursor => _inputWithFauxCursor;
        public string Input
        {
            set
            {
                _inputWithFauxCursor = value.Contains(RenameTests.FAUX_CURSOR) ? value : _inputWithFauxCursor;
                _codeString = value.ToCodeString();
            }
            get
            {
                return _codeString.Code;
            }
        }

        public string Expected
        {
            set
            {
                _expected = value;
            }
            get
            {
                return _expected.Equals(string.Empty) ? Input : _expected;
            }
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
            _inputWithFauxCursor = string.Empty;
            _expected = string.Empty;
            ModuleName = moduleName;
            ModuleType = moduleType;
            CheckExpectedEqualsActual = true;
            ControlNames = new List<String>();
            NewName = "";
        }
    }
}
