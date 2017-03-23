using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using VbaBlocks = RubberduckTests.Inspections.ObjectVariableNotSetInspectionTests_VBABlocks;


namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObjectVariableNotSetInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenIndexerObjectAccess_ReturnsNoResult()
        {
            var tp = VbaBlocks.GivenIndexerObjectAccess_ReturnsNoResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenIndexerObjectAccess_ReturnsResult()
        {
            var tp = VbaBlocks.GivenIndexerObjectAccess_ReturnsResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenStringVariable_ReturnsNoResult()
        {
            var tp = VbaBlocks.GivenStringVariable_ReturnsNoResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedObject_ReturnsResult()
        {
            var tp = VbaBlocks.GivenVariantVariableAssignedObject_ReturnsResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedNewObject_ReturnsResult()
        {
            var tp = VbaBlocks.GivenVariantVariableAssignedNewObject_ReturnsResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]//todo
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedRange_ReturnsResult()
        {
            var tp = VbaBlocks.GivenVariantVariableAssignedRange_ReturnsResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedDeclaredRange_ReturnsResult()
        {
            var tp = VbaBlocks.GivenVariantVariableAssignedDeclaredRange_ReturnsResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedDeclaredVariant_ReturnsNoResult()
        {
            var tp = VbaBlocks.GivenVariantVariableAssignedDeclaredVariant_ReturnsNoResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedBaseType_ReturnsNoResult()
        {
            var tp = VbaBlocks.GivenVariantVariableAssignedBaseType_ReturnsNoResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenObjectVariableNotSet_ReturnsResult()
        {
            var tp = VbaBlocks.GivenObjectVariableNotSet_ReturnsResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenObjectVariableNotSet_Ignored_DoesNotReturnResult()
        {
            var tp = VbaBlocks.GivenObjectVariableNotSet_Ignored_DoesNotReturnResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenSetObjectVariable_ReturnsNoResult()
        {
            var tp = VbaBlocks.GivenSetObjectVariable_ReturnsNoResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2266
        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_FunctionReturnsArrayOfType_ReturnsNoResult()
        {
            var testParams = VbaBlocks.FunctionReturnsArrayOfType_ReturnsNoResult_TestParams();

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Codez", ComponentType.StandardModule, testParams.Key)
                .AddReference("Scripting", "", 1, 0, true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.State.AddTestLibrary("Scripting.1.0.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObjectVariableNotSetInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(testParams.Value, inspectionResults.Count());

        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_IgnoreQuickFixWorks()
        {
            var inputCode = VbaBlocks.IgnoreQuickFixWorks_Input();
            var expectedCode = VbaBlocks.IgnoreQuickFixWorks_Expected();

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObjectVariableNotSetInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_ForFunctionAssignment_ReturnsResult()
        {
            var testParams = VbaBlocks.ForFunctionAssignment_ReturnsResult_TestParams();
            var expectedCode = VbaBlocks.ForFunctionAssignment_ReturnsResult_Expected();

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(testParams.Key, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObjectVariableNotSetInspection(state);
            var inspectionResults = inspection.GetInspectionResults().ToList();

            Assert.AreEqual(testParams.Value, inspectionResults.Count);
            foreach (var fix in inspectionResults.SelectMany(result => result.QuickFixes.Where(s => s is UseSetKeywordForObjectAssignmentQuickFix)))
            {
                fix.Fix();
            }
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_ForPropertyGetAssignment_ReturnsResults()
        {
            var testParams = VbaBlocks.ForPropertyGetAssignment_ReturnsResults_TestParams();
            var expectedCode = VbaBlocks.ForPropertyGetAssignment_ReturnsResults_Expected();

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(testParams.Key, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObjectVariableNotSetInspection(state);
            var inspectionResults = inspection.GetInspectionResults().ToList();

            Assert.AreEqual(testParams.Value, inspectionResults.Count);
            foreach (var fix in inspectionResults.SelectMany(result => result.QuickFixes.Where(s => s is UseSetKeywordForObjectAssignmentQuickFix)))
            {
                fix.Fix();
            }
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_LongPtrVariable_ReturnsNoResult()
        {
            var tp = VbaBlocks.LongPtrVariable_ReturnsNoResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key,tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_NoTypeSpecified_ReturnsResult()
        {
            var tp = VbaBlocks.NoTypeSpecified_ReturnsResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_SelfAssigned_ReturnsNoResult()
        {
            var tp = VbaBlocks.SelfAssigned_ReturnsNoResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_EnumVariable_ReturnsNoResult()
        {

            var tp = VbaBlocks.EnumVariable_ReturnsNoResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_FunctionReturnNotSet_ReturnsResult()
        {

            var tp = VbaBlocks.FunctionReturnNotSet_ReturnsResult_TestParams();
            AssertInputCodeYieldsExpectedInspectionResultCount(tp.Key, tp.Value);
        }

        private void AssertInputCodeYieldsExpectedInspectionResultCount(string inputCode, int expected)
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObjectVariableNotSetInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(expected, inspectionResults.Count());
        }
    }
}

