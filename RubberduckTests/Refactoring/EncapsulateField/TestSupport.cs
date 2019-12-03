using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    public class EncapsulateFieldTestSupport : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private TestEncapsulationAttributes UserModifiedEncapsulationAttributes(string field, string property = null, bool? isReadonly = null, bool encapsulateFlag = true)
        {
            var testAttrs = new TestEncapsulationAttributes(field, encapsulateFlag, isReadonly ?? false);
            if (property != null)
            {
                testAttrs.PropertyName = property;
            }
            return testAttrs;
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> UserAcceptsDefaults(bool asUDT = false)
        {
            return model => 
            {
                model.EncapsulateWithUDT = asUDT;
                return model;
            };
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParametersForSingleTarget(string field, string property = null, bool? isReadonly = null, bool encapsulateFlag = true, bool asUDT = false)
        {
            var clientAttrs = UserModifiedEncapsulationAttributes(field, property, isReadonly, encapsulateFlag);

            return SetParameters(field, clientAttrs, asUDT);
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(UserInputDataObject userModifications)
        {
            return model =>
            {
                model.EncapsulateWithUDT = userModifications.EncapsulateAsUDT;
                foreach (var testModifiedAttribute in userModifications.EncapsulateFieldAttributes)
                {
                    var attrsInitializedByTheRefactoring = model[testModifiedAttribute.TargetFieldName].EncapsulationAttributes;

                    attrsInitializedByTheRefactoring.PropertyName = testModifiedAttribute.PropertyName;
                    attrsInitializedByTheRefactoring.EncapsulateFlag = testModifiedAttribute.EncapsulateFlag;

                    var currentAttributes = model[testModifiedAttribute.TargetFieldName].EncapsulationAttributes;
                    currentAttributes.PropertyName = attrsInitializedByTheRefactoring.PropertyName;
                    currentAttributes.EncapsulateFlag = attrsInitializedByTheRefactoring.EncapsulateFlag;

                }
                foreach ((string instanceVariable, string memberName, bool flag) in userModifications.UDTMemberNameFlagPairs)
                {
                    model[$"{instanceVariable}.{memberName}"].EncapsulateFlag = flag;
                }
                return model;
            };
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(string originalField, TestEncapsulationAttributes attrs, bool asUDT = false)
        {
            return model =>
            {
                var encapsulatedField = model[originalField];
                encapsulatedField.EncapsulationAttributes.PropertyName = attrs.PropertyName;
                encapsulatedField.EncapsulationAttributes.IsReadOnly = attrs.IsReadOnly;
                encapsulatedField.EncapsulationAttributes.EncapsulateFlag = attrs.EncapsulateFlag;

                model.EncapsulateWithUDT = asUDT;
                return model;
            };
        }

        public string RefactoredCode(CodeString codeString, Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAdjustment, Type expectedException = null, bool executeViaActiveSelection = false)
            => RefactoredCode(codeString.Code, codeString.CaretPosition.ToOneBased(), presenterAdjustment, expectedException, executeViaActiveSelection);

        public IRefactoring SupportTestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            var indenter = CreateIndenter(); //The refactoring only uses method independent of the VBE instance.
            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            return new EncapsulateFieldRefactoring(state, indenter, factory, rewritingManager, selectionService, selectedDeclarationProvider);
        }

        private static IIndenter CreateIndenter(IVBE vbe = null)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IRefactoringPresenterFactory factory, ISelectionService selectionService)
        {
            return SupportTestRefactoring(rewritingManager, state, factory, selectionService);
        }
    }

    public class TestEncapsulationAttributes
    {
        public TestEncapsulationAttributes(string fieldName, bool encapsulationFlag = true, bool isReadOnly = false)
        {
            _identifiers = new EncapsulationIdentifiers(fieldName);
            EncapsulateFlag = encapsulationFlag;
            IsReadOnly = isReadOnly;
        }

        private EncapsulationIdentifiers _identifiers;
        public string TargetFieldName => _identifiers.TargetFieldName;

        public string NewFieldName
        {
            get => _identifiers.Field;
            set => _identifiers.Field = value;
        }
        public string PropertyName
        {
            get => _identifiers.Property;
            set => _identifiers.Property = value;
        }
        public bool EncapsulateFlag { get; set; }
        public bool IsReadOnly { get; set; }
    }

    public class UserInputDataObject
    {
        private List<TestEncapsulationAttributes> _userInput = new List<TestEncapsulationAttributes>();
        private List<(string, string, bool)> _udtNameFlagPairs = new List<(string, string, bool)>();

        public UserInputDataObject(string fieldName, string propertyName = null, bool encapsulationFlag = true, bool isReadOnly = false)
            => AddAttributeSet(fieldName, propertyName, encapsulationFlag, isReadOnly);

        public void AddAttributeSet(string fieldName, string propertyName = null, bool encapsulationFlag = true, bool isReadOnly = false)
        {
            var attrs = new TestEncapsulationAttributes(fieldName, encapsulationFlag, isReadOnly);
            attrs.PropertyName = propertyName ?? attrs.PropertyName;

            _userInput.Add(attrs);
        }

        public bool EncapsulateAsUDT { set; get; }

        public TestEncapsulationAttributes this[string fieldName]
            => EncapsulateFieldAttributes.Where(efa => efa.TargetFieldName == fieldName).Single();


        public IEnumerable<TestEncapsulationAttributes> EncapsulateFieldAttributes => _userInput;

        public void AddUDTMemberNameFlagPairs(params (string, string, bool)[] nameFlagPairs)
            => _udtNameFlagPairs.AddRange(nameFlagPairs);

        public IEnumerable<(string, string, bool)> UDTMemberNameFlagPairs => _udtNameFlagPairs;
    }
}
