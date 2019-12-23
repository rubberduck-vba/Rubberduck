using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    public class EncapsulateFieldTestSupport : InteractiveRefactoringTestBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        public string StateUDTDefaultType => $"T{MockVbeBuilder.TestModuleName}";

        private TestEncapsulationAttributes UserModifiedEncapsulationAttributes(string field, string property = null, bool isReadonly = false, bool encapsulateFlag = true)
        {
            var testAttrs = new TestEncapsulationAttributes(field, encapsulateFlag, isReadonly);
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

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> UserAcceptsDefaults(params string[] fieldNames)
        {
            return model =>
            {
                foreach (var name in fieldNames)
                {
                    model[name].EncapsulateFlag = true;
                }
                return model;
            };
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParametersForSingleTarget(string field, string property = null, bool isReadonly = false, bool encapsulateFlag = true, bool asUDT = false)
        {
            var clientAttrs = UserModifiedEncapsulationAttributes(field, property, isReadonly, encapsulateFlag);

            return SetParameters(field, clientAttrs, asUDT);
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(UserInputDataObject userInput)
        {
            return model =>
            {
                model.EncapsulateWithUDT = userInput.EncapsulateAsUDT;
                if (userInput.EncapsulateAsUDT)
                {
                    var stateUDT = model.SelectedFieldCandidates.Where(sfc => sfc is IUserDefinedTypeCandidate udt && udt.TargetID == userInput.ObjectStateUDTTargetID)
                    .Select(sfc => sfc as IUserDefinedTypeCandidate).SingleOrDefault();
                    if (stateUDT != null)
                    {
                        stateUDT.IsObjectStateUDT = userInput.ObjectStateUDTTargetID != null;
                        model.StateUDTField = new ObjectStateUDT(stateUDT);
                    }
                }

                foreach (var testModifiedAttribute in userInput.EncapsulateFieldAttributes)
                {
                    var attrsInitializedByTheRefactoring = model[testModifiedAttribute.TargetFieldName]; //.EncapsulationAttributes;

                    attrsInitializedByTheRefactoring.PropertyName = testModifiedAttribute.PropertyName;
                    attrsInitializedByTheRefactoring.EncapsulateFlag = testModifiedAttribute.EncapsulateFlag;
                    attrsInitializedByTheRefactoring.IsReadOnly = testModifiedAttribute.IsReadOnly;
                }
                return model;
            };
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(string originalField, TestEncapsulationAttributes attrs, bool asUDT = false)
        {
            return model =>
            {
                var encapsulatedField = model[originalField];
                encapsulatedField.PropertyName = attrs.PropertyName;
                encapsulatedField.IsReadOnly = attrs.IsReadOnly;
                encapsulatedField.EncapsulateFlag = attrs.EncapsulateFlag;

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
            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());
            return new EncapsulateFieldRefactoring(state, indenter, factory, rewritingManager, selectionService, selectedDeclarationProvider, uiDispatcherMock.Object);
        }

        public IEncapsulateFieldCandidate RetrieveEncapsulatedField(string inputCode, string fieldName)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;

            var selectedComponentName = vbe.SelectedVBComponent.Name;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var match = state.DeclarationFinder.MatchName(fieldName).Single();
                return new EncapsulateFieldCandidate(match, new EncapsulateFieldNamesValidator(state)) as IEncapsulateFieldCandidate;
            }
        }

        public IEncapsulateFieldCandidate RetrieveEncapsulatedField(string inputCode, string fieldName, DeclarationType declarationType)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;

            var selectedComponentName = vbe.SelectedVBComponent.Name;

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var match = state.DeclarationFinder.MatchName(fieldName).Where(m => m.DeclarationType.Equals(declarationType)).Single();
                return new EncapsulateFieldCandidate(match, new EncapsulateFieldNamesValidator(state)) as IEncapsulateFieldCandidate;
            }
        }

        public EncapsulateFieldModel RetrieveUserModifiedModelPriorToRefactoring(IVBE vbe, string declarationName, DeclarationType declarationType, Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAdjustment) //, params string[] fieldIdentifiers)
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var targets = state.DeclarationFinder.DeclarationsWithType(declarationType);

                var target = targets.Single(declaration => declaration.IdentifierName == declarationName);

                var refactoring = TestRefactoring(rewritingManager, state, presenterAdjustment);
                if (refactoring is IEncapsulateFieldRefactoringTestAccess concrete)
                {
                    return concrete.TestUserInteractionOnly(target, presenterAdjustment);
                }
                throw new InvalidCastException();
            }
        }

        public static IIndenter CreateIndenter(IVBE vbe = null)
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

        public UserInputDataObject() { }

        //public UserInputDataObject(string fieldName, string propertyName = null, /*bool encapsulationFlag = true,*/ bool isReadOnly = false)
        //    : this()
        //{
        //    UserSelectsField(fieldName, propertyName/*, encapsulationFlag*/, isReadOnly);
        //}

        public UserInputDataObject UserSelectsField(string fieldName, string propertyName = null/*, bool encapsulationFlag = true*/, bool isReadOnly = false)
        {
            //var attrs = new TestEncapsulationAttributes(fieldName, true, isReadOnly);
            //attrs.PropertyName = propertyName ?? attrs.PropertyName;
            //attrs.EncapsulateFlag = true;
            //attrs.IsReadOnly = isReadOnly;

            //_userInput.Add(attrs);
            //return this;
            return AddUserInputSet(fieldName, propertyName, true, isReadOnly);
        }

        public UserInputDataObject AddUserInputSet(string fieldName, string propertyName = null, bool encapsulationFlag = true, bool isReadOnly = false)
        {
            var attrs = new TestEncapsulationAttributes(fieldName, encapsulationFlag, isReadOnly);
            attrs.PropertyName = propertyName ?? attrs.PropertyName;
            attrs.EncapsulateFlag = encapsulationFlag;
            attrs.IsReadOnly = isReadOnly;

            _userInput.Add(attrs);
            return this;
        }

        public bool EncapsulateAsUDT { set; get; }

        public string ObjectStateUDTTargetID { set; get; }

        public string StateUDT_TypeName { set; get; }

        public string StateUDT_FieldName { set; get; }

        public TestEncapsulationAttributes this[string fieldName]
            => EncapsulateFieldAttributes.Where(efa => efa.TargetFieldName == fieldName).Single();

        public IEnumerable<TestEncapsulationAttributes> EncapsulateFieldAttributes => _userInput;

        //public void AddUDTMemberNameFlagPairs(params (string, string, bool)[] nameFlagPairs)
        //    => _udtNameFlagPairs.AddRange(nameFlagPairs);

        //public IEnumerable<(string, string, bool)> UDTMemberNameFlagPairs => _udtNameFlagPairs;
    }
}
