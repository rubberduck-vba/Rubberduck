using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.Rename
{
    internal static class RenameTestExecution
    {

        public static void PerformExpectedVersusActualRenameTests(RenameTestsDataObject tdo
            , RenameTestModuleDefinition? inputOutput1
            , RenameTestModuleDefinition? inputOutput2 = null
            , RenameTestModuleDefinition? inputOutput3 = null
            , RenameTestModuleDefinition? inputOutput4 = null)
        {
            try
            {
                InitializeTestDataObject(tdo, inputOutput1, inputOutput2, inputOutput3, inputOutput4);
                RunRenameRefactorScenario(tdo);
                CheckRenameRefactorTestResults(tdo);
            }
            finally
            {
                tdo.ParserState?.Dispose();
            }
        }

        private static void InitializeTestDataObject(RenameTestsDataObject tdo
            , RenameTestModuleDefinition? inputOutput1
            , RenameTestModuleDefinition? inputOutput2 = null
            , RenameTestModuleDefinition? inputOutput3 = null
            , RenameTestModuleDefinition? inputOutput4 = null)
        {
            var renameTMDs = new List<RenameTestModuleDefinition>();
            bool cursorFound = false;
            foreach (var io in new[] { inputOutput1, inputOutput2, inputOutput3, inputOutput4 })
            {
                if (io.HasValue)
                {
                    var renameTMD = io.Value;
                    if (!renameTMD.Input_WithFauxCursor.Equals(string.Empty))
                    {
                        if (cursorFound) { Assert.Inconclusive($"Found multiple selection cursors ('{RenameTests.FAUX_CURSOR}') in the test input"); }
                        cursorFound = true;
                    }
                    renameTMDs.Add(renameTMD);
                }
            }

            if (!cursorFound)
            {
                Assert.Inconclusive($"Unable to determine selected target using '{RenameTests.FAUX_CURSOR}' in test input");
            }

            renameTMDs.ForEach(rtmd => AddTestModuleDefinition(tdo, rtmd));

            if (tdo.NewName.Length == 0)
            {
                Assert.Inconclusive("NewName is blank");
            }
            if (!tdo.RawSelection.HasValue)
            {
                Assert.Inconclusive("A User 'Selection' has not been defined for the test");
            }

            tdo.MsgBox = new Mock<IMessageBox>();
            tdo.MsgBox.Setup(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>())).Returns(tdo.MsgBoxReturn == ConfirmationOutcome.Yes);
            
            var activeIndex = renameTMDs.FindIndex(tmd => tmd.Input_WithFauxCursor != string.Empty);
            tdo.VBE = tdo.VBE ?? BuildProject(tdo.ProjectName, tdo.ModuleTestSetupDefs, activeIndex, tdo.UseLibraries);
            tdo.AdditionalSetup?.Invoke(tdo);
            (tdo.ParserState, tdo.RewritingManager) = MockParser.CreateAndParseWithRewritingManager(tdo.VBE);

            CreateQualifiedSelectionForTestCase(tdo);
            tdo.RenameModel = new RenameModel(tdo.ParserState.DeclarationFinder, tdo.QualifiedSelection) { NewName = tdo.NewName };
            Assert.IsTrue(tdo.RenameModel.Target.IdentifierName.Contains(tdo.OriginalName)
                , $"Target aquired ({tdo.RenameModel.Target.IdentifierName} does not equal name specified ({tdo.OriginalName}) in the test");

            var presenter = new Mock<IRenamePresenter>();
            var factory = GetFactoryMock(m => {
                presenter.Setup(p => p.Model).Returns(m);
                presenter.Setup(p => p.Show(It.IsAny<Declaration>()))
                    .Callback(() => m.NewName = tdo.NewName)
                    .Returns(m);
                presenter.Setup(p => p.Show())
                    .Callback(() => m.NewName = tdo.NewName)
                    .Returns(m);
                return presenter;
            }, out var creator);

            tdo.RenameRefactoringUnderTest = new RenameRefactoring(tdo.VBE, factory.Object, tdo.MsgBox.Object, tdo.ParserState, tdo.ParserState.ProjectsProvider, tdo.RewritingManager);
        }

        private static void AddTestModuleDefinition(RenameTestsDataObject tdo, RenameTestModuleDefinition inputOutput)
        {
            if (inputOutput.Input_WithFauxCursor.Length > 0)
            {
                tdo.SelectionModuleName = inputOutput.ModuleName;
                if (inputOutput.Input_WithFauxCursor.Contains(RenameTests.FAUX_CURSOR))
                {
                    var numCursors = inputOutput.Input_WithFauxCursor.ToArray().Where(c => c.Equals(RenameTests.FAUX_CURSOR)).Count();
                    if (numCursors != 1)
                    {
                        Assert.Inconclusive($"{numCursors} found in FauxCursor input - only a single cursor is allowed.");
                    }
                    tdo.RawSelection = inputOutput.RenameSelection;
                    if (!tdo.RawSelection.HasValue)
                    {
                        Assert.Inconclusive($"Unable to set RawSelection field for test module {inputOutput.ModuleName}");
                    }

                    //FIXME is this still necessary? I think not...
                    //inputOutput.RenameSelection = tdo.RawSelection;
                }
            }
            tdo.ModuleTestSetupDefs.Add(inputOutput);
        }

        private static void RunRenameRefactorScenario(RenameTestsDataObject tdo)
        {
            if (tdo.RefactorParamType == RefactorParams.Declaration)
            {
                tdo.RenameRefactoringUnderTest.Refactor(tdo.RenameModel.Target);
            }
            else if (tdo.RefactorParamType == RefactorParams.QualifiedSelection)
            {
                tdo.RenameRefactoringUnderTest.Refactor(tdo.QualifiedSelection);
            }
            else
            {
                tdo.RenameRefactoringUnderTest.Refactor();
            }
        }

        private static void CheckRenameRefactorTestResults(RenameTestsDataObject tdo)
        {
            foreach (var inputOutput in tdo.ModuleTestSetupDefs)
            {
                if (inputOutput.CheckExpectedEqualsActual)
                {
                    var codeModule = RetrieveComponent(tdo, inputOutput.ModuleName).CodeModule;
                    var expected = inputOutput.Expected;
                    var actual = codeModule.Content();
                    Assert.AreEqual(expected, actual);
                }
            }
        }

        public static Mock<IRefactoringPresenterFactory> GetFactoryMock(Func<RenameModel, Mock<IRenamePresenter>> mockedPresenter, out RememberingCreator<RenameModel, IRenamePresenter> creator)
        {
            var m = new Mock<IRefactoringPresenterFactory>();
            var c = new RememberingCreator<RenameModel, IRenamePresenter>(mockedPresenter);
            m.Setup(f => f.Create<IRenamePresenter, RenameModel>(It.IsAny<RenameModel>()))
                .Returns<RenameModel>(input => c.DoCreate(input).Object);
            m.Setup(f => f.Release(It.Is<IRenamePresenter>(p => p == c.Memento.Object)));
            creator = c;
            return m;
        }

        public class RememberingCreator<I, O>
            where I : class
            where O : class
        {
            private readonly Func<I, Mock<O>> _delegate;

            public RememberingCreator(Func<I, Mock<O>> d)
            {
                _delegate = d;
            }

            public Mock<O> Memento { get; private set; }
            public Mock<O> DoCreate(I input)
            {
                var mock = _delegate.Invoke(input);
                Memento = mock;
                return mock;
            }
        }

        private static void CreateQualifiedSelectionForTestCase(RenameTestsDataObject tdo)
        {
            var component = RetrieveComponent(tdo, tdo.SelectionModuleName);
            if (tdo.RawSelection.HasValue)
            {
                tdo.QualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), tdo.RawSelection.Value);
                return;
            }
            Assert.Inconclusive($"Unable to find target '{RenameTests.FAUX_CURSOR}' in { tdo.SelectionModuleName} content.");
        }

        private static IVBE BuildProject(string projectName, List<RenameTestModuleDefinition> testComponents, int activeIndex, bool useLibraries = false)
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected);

            if (useLibraries)
            {
                enclosingProjectBuilder.AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 1, true);
                enclosingProjectBuilder.AddReference("EXCEL", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true);
            }

            foreach (var comp in testComponents)
            {
                if (comp.ModuleType == ComponentType.UserForm)
                {
                    var form = enclosingProjectBuilder.MockUserFormBuilder(comp.ModuleName, comp.Input);
                    if (!comp.ControlNames.Any())
                    {
                        Assert.Inconclusive("Test incorporates a UserForm but does not define any controls");
                    }
                    foreach (var control in comp.ControlNames)
                    {
                        form.AddControl(control);
                    }
                    (var component, var codeModule) = form.Build();
                    enclosingProjectBuilder.AddComponent(component, codeModule);
                }
                else
                {
                    var selection = comp.RenameSelection.HasValue ? comp.RenameSelection.Value : default;
                    enclosingProjectBuilder.AddComponent(comp.ModuleName, comp.ModuleType, comp.Input, selection);
                }
            }
            var project = enclosingProjectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();
            vbe.SetupGet(v => v.ActiveCodePane).Returns(project.Object.VBComponents[activeIndex].CodeModule.CodePane);
            return vbe.Object;
        }

        internal static IVBComponent RetrieveComponent(RenameTestsDataObject tdo, string componentName)
        {
            var vbProject = tdo.VBE.VBProjects.Single(item => item.Name == tdo.ProjectName);
            return vbProject.VBComponents.SingleOrDefault(item => item.Name == componentName);
        }

        internal enum RefactorParams
        {
            None,
            QualifiedSelection,
            Declaration
        };
    }
}
