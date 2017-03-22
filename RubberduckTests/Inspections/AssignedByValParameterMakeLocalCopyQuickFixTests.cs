using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.UI.Refactorings;
using System.Windows.Forms;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using VbaCodeBlocks = RubberduckTests.Inspections.AssignedByValParameterMakeLocalCopyQuickFixTests_VBABlocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class AssignedByValParameterMakeLocalCopyQuickFixTests
    {

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment()
        {
            var inputCode = VbaCodeBlocks.LocalVariableAssignment_Input();
            var expectedCode = VbaCodeBlocks.LocalVariableAssignment_Expected();

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        //weaponized formatting
        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_ComplexFormat()
        {
            var inputCode = VbaCodeBlocks.LocalVariableAssignment_ComplexFormat_Input();
            var expectedCode = VbaCodeBlocks.LocalVariableAssignment_ComplexFormat_Expected();

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_ComputedNameAvoidsCollision()
        {
            var inputCode = VbaCodeBlocks.LocalVariableAssignment_ComputedNameAvoidsCollision_Input();
            var expectedCode = VbaCodeBlocks.LocalVariableAssignment_ComputedNameAvoidsCollision_Expected();

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUseOtherSub()
        {
            //Make sure the modified code stays within the specific method under repair
            var inputCode = VbaCodeBlocks.LocalVariableAssignment_NameInUseOtherSub_Input();
            var expectedFragment = VbaCodeBlocks.LocalVariableAssignment_NameInUseOtherSub_Expected();

            string[] splitToken = { "'VerifyNoChangeBelowThisLine" };
            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            var evaluatedFragment = quickFixResult.Split(splitToken, System.StringSplitOptions.None)[1];
            Assert.AreEqual(expectedFragment, evaluatedFragment);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUseOtherProperty()
        {
            //Make sure the modified code stays within the specific method under repair
            var inputCode = VbaCodeBlocks.LocalVariableAssignment_NameInUseOtherProperty_Input();
            var expectedCode = VbaCodeBlocks.LocalVariableAssignment_NameInUseOtherProperty_Expected();

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            string[] splitToken = VbaCodeBlocks.LocalVariable_NameInUseOtherProperty_SplitToken();
            var evaluatedResult = quickFixResult.Split(splitToken, System.StringSplitOptions.None)[1];

            Assert.AreEqual(expectedCode, evaluatedResult);
        }

        //Replicates issue #2873 : AssignedByValParameter quick fix needs to use `Set` for reference types.
        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_UsesSet()
        {
            var inputCode = VbaCodeBlocks.LocalVariableAssignment_UsesSet_Input();
            var expectedCode = VbaCodeBlocks.LocalVariableAssignment_UsesSet_Expected();

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NoAsTypeClause()
        {
            var inputCode = VbaCodeBlocks.LocalVariableAssignment_NoAsTypeClause_Input();
            var expectedCode = VbaCodeBlocks.LocalVariableAssignment_NoAsTypeClause_Expected();

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_EnumType()
        {
            var inputCode = VbaCodeBlocks.LocalVariableAssignment_EnumType_Input();
            var expectedCode = VbaCodeBlocks.LocalVariableAssignment_EnumType_Expected();

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        private string ApplyLocalVariableQuickFixToCodeFragment(string inputCode, string userEnteredName = "")
        {
            var vbe = BuildMockVBE(inputCode);

            var mockDialogFactory = BuildMockDialogFactory(userEnteredName);

            RubberduckParserState state;
            var inspectionResults = GetInspectionResults(vbe.Object, mockDialogFactory.Object, out state);
            var result = inspectionResults.FirstOrDefault();
            if (result == null)
            {
                Assert.Inconclusive("Inspection yielded no results.");
            }

            result.QuickFixes.Single(s => s is AssignedByValParameterMakeLocalCopyQuickFix).Fix();
            return state.GetRewriter(vbe.Object.ActiveVBProject.VBComponents[0]).GetText();
        }

        private Mock<IVBE> BuildMockVBE(string inputCode)
        {
            IVBComponent component;
            return MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
        }

        private IEnumerable<IInspectionResult> GetInspectionResults(IVBE vbe, IAssignedByValParameterQuickFixDialogFactory mockDialogFactory, out RubberduckParserState state)
        {
            state = MockParser.CreateAndParse(vbe);

            var inspection = new AssignedByValParameterInspection(state, mockDialogFactory);
            return inspection.GetInspectionResults();
        }

        private Mock<IAssignedByValParameterQuickFixDialogFactory> BuildMockDialogFactory(string userEnteredName)
        {
            var mockDialog = new Mock<IAssignedByValParameterQuickFixDialog>();

            mockDialog.SetupAllProperties();

            if (userEnteredName.Length > 0)
            {
                mockDialog.SetupGet(m => m.NewName).Returns(() => userEnteredName);
            }
            mockDialog.SetupGet(m => m.DialogResult).Returns(() => DialogResult.OK);

            var mockDialogFactory = new Mock<IAssignedByValParameterQuickFixDialogFactory>();
            mockDialogFactory.Setup(f => f.Create(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IEnumerable<string>>())).Returns(mockDialog.Object);

            return mockDialogFactory;
        }
    }
}
