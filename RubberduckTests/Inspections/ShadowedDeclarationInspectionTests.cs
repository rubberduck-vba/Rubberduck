using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ShadowedDeclarationInspectionTests
    {
        private const string ProjectName = "SameNameProject";
        private const string ProceduralModuleName = "SameNameProceduralModule";
        private const string ClassModuleName = "SameNameClass";
        private const string UserFormName = "SameNameUserForm";
        private const string DocumentName = "SameNameDocument";
        private const string ProcedureName = "SameNameProcedure";
        private const string FunctionName = "SameNameFunction";
        private const string PropertyGetName = "SameNamePropertyGet";
        private const string PropertySetName = "SameNamePropertySet";
        private const string PropertyLetName = "SameNamePropertyLet";
        private const string ParameterName = "SameNameParameter";
        private const string VariableName = "SameNameVariable";
        private const string LocalVariableName = "SameNameLocalVariable";
        private const string ConstantName = "SameNameConstant";
        private const string LocalConstantName = "SameNameLocalConstant";
        private const string EnumerationName = "SameNameEnumeration";
        private const string EnumerationMemberName = "SameNameEnumerationMember";
        private const string EventName = "SameNameEvent";
        private const string UserDefinedTypeName = "SameNameUserDefinedType";
        private const string UserDefinedTypeMemberName = "SameNameUserDefinedTypeMember";
        private const string LibraryProcedureName = "SameNameLibraryProcedure";
        private const string LibraryFunctionName = "SameNameLibraryFunction";
        private const string LineLabelName = "SameNameLineLabel";

        private readonly string _moduleCode =
            $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
Dim {LocalVariableName} as String
Const {LocalConstantName} as String = """"
{LineLabelName}:
End Sub

Public Function {FunctionName}()
End Function

Public Property Get {PropertyGetName}()
End Property

Public Property Let {PropertySetName}(v As Variant)
End Property

Public Property Set {PropertyLetName}(s As String)
End Property";

        private readonly string _classCode =
            $@"Public {VariableName} As String

Public Event {EventName}()

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
{LineLabelName}:
End Sub

Public Function {FunctionName}()
End Function

Public Property Get {PropertyGetName}()
End Property

Public Property Let {PropertySetName}(v As Variant)
End Property

Public Property Set {PropertyLetName}(s As String)
End Property";

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };


            var builder = new MockVbeBuilder();
            var userProjectBuilder = CreateUserProject(builder);
            foreach (var expectedResultCount in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedProject = builder.ProjectBuilder(expectedResultCount.Key, ProjectProtection.Unprotected)
                    .AddComponent("Foo" + expectedResultCount.Key, ComponentType.StandardModule, "")
                    .Build();
                builder.AddProject(referencedProject);
                userProjectBuilder = userProjectBuilder.AddReference(expectedResultCount.Key, "");
            }

            var userProject = userProjectBuilder.Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            foreach (var expectedResultCount in expectedResultCountsByDeclarationIdentifierName)
            {
                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder, expectedResultCount.Key).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();

                IEnumerable<IInspectionResult> inspectionResults;
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {
                    var inspection = new ShadowedDeclarationInspection(state);
                    inspectionResults = inspection.GetInspectionResults();
                }

                Assert.AreEqual(expectedResultCount.Value, inspectionResults.Count(), $"Wrong number of inspection results for {expectedResultCount.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsProceduralModuleInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentsOfOneType(
                referencedProjectName: "Foo",
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                referencedComponentsComponentType: ComponentType.StandardModule,
                componentNameSelector: key => key, 
                componentCodeSelector: key => string.Empty);
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsProceduralModuleInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentsOfOneType(
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                additionalComponentsComponentType: ComponentType.StandardModule,
                componentNameSelector: key => key,
                componentCodeSelector: key => string.Empty);
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsNonExposedClassModuleInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentsOfOneType(
                referencedProjectName: "Foo",
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                referencedComponentsComponentType: ComponentType.ClassModule,
                componentNameSelector: key => key,
                componentCodeSelector: key => string.Empty);
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserFormInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentsOfOneType(
                referencedProjectName: "Foo",
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                referencedComponentsComponentType: ComponentType.UserForm,
                componentNameSelector: key => key,
                componentCodeSelector: key => string.Empty);
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserFormInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentsOfOneType(
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                additionalComponentsComponentType: ComponentType.UserForm,
                componentNameSelector: key => key,
                componentCodeSelector: key => string.Empty);
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsDocumentInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentsOfOneType(
                referencedProjectName: "Foo",
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                referencedComponentsComponentType: ComponentType.Document,
                componentNameSelector: key => key,
                componentCodeSelector: key => string.Empty);
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsDocumentInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentsOfOneType(
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                additionalComponentsComponentType: ComponentType.Document,
                componentNameSelector: key => key,
                componentCodeSelector: key => string.Empty);
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicProcedureInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub {key}()
End Sub");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateProcedureInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Sub {key}()
End Sub");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicProcedureInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub {key}()
End Sub");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateProcedureInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Sub {key}()
End Sub");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsProcedureInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

                var baseCode =
                    $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Function {FunctionName}({ParameterName} As String)
    Dim {LocalVariableName} as String
    Const {LocalConstantName} as String = """"
{LineLabelName}:
End Function

Public Property Get {PropertyGetName}()
End Property

Public Property Let {PropertySetName}(v As Variant)
End Property

Public Property Set {PropertyLetName}(s As String)
End Property";

                var code = ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(
                    baseCode: baseCode,
                    testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                    moduleBodyElementCodeSelector: key => $@"Public Sub {key}()
End Sub");
                var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicFunctionInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Function {key}()
End Function");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateFunctionInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Function {key}()
End Function");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicFunctionInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Function {key}()
End Function");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateFunctionInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Function {key}()
End Function");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsFunctionInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };


            var baseCode =
                $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
Dim {LocalVariableName} as String
Const {LocalConstantName} as String = """"
{LineLabelName}:
End Sub

Public Property Get {PropertyGetName}()
End Property

Public Property Let {PropertySetName}(v As Variant)
End Property

Public Property Set {PropertyLetName}(s As String)
End Property";

            var code = ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Function {key}()
End Function");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicPropertyGetInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Property Get {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivatePropertyGetInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Property Get {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicPropertyGetInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Property Get {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivatePropertyGetInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Property Get {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPropertyGetInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };


            var baseCode =
                $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
Dim {LocalVariableName} as String
Const {LocalConstantName} as String = """"
{LineLabelName}:
End Sub

Public Function {FunctionName}()
End Function

Public Property Let {PropertySetName}(v As Variant)
End Property

Public Property Set {PropertyLetName}(s As String)
End Property";

            var code = ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Property Get {key}()
End Property");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicPropertySetInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Property Set {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivatePropertySetInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Property Set {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicPropertySetInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Property Set {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivatePropertySetInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Property Set {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPropertySetInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var baseCode =
                $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
Dim {LocalVariableName} as String
Const {LocalConstantName} as String = """"
{LineLabelName}:
End Sub

Public Function {FunctionName}()
End Function

Public Property Get {PropertyGetName}()
End Property

Public Property Let {PropertySetName}(v As Variant)
End Property";

            var code = ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Property Set {key}()
End Property");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicPropertyLetInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Property Let {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivatePropertyLetInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Property Let {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicPropertyLetInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Property Let {key}()
End Property");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivatePropertyLetInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Private Property Let {key}()
End Property"); 
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPropertyLetInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var baseCode =
                $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
Dim {LocalVariableName} as String
Const {LocalConstantName} as String = """"
{LineLabelName}:
End Sub

Public Function {FunctionName}()
End Function

Public Property Get {PropertyGetName}()
End Property

Public Property Set {PropertySetName}(v As Variant)
End Property";

            var code = ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Property Let {key}()
End Property");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsParameterInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub Qux{key}({key} As String)
End Sub");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsParameterInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentsComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub Qux{key}({key} As String)
End Sub");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsParameterInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            var code = ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(
                baseCode: _moduleCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Function Foo{key}({key} As String)
End Function");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsVariableInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Global
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Global {result.Key} As String";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for global {result.Key}");
                }
            }

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Public {result.Key} As String";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Private {result.Key} As String";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(0, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsVariableInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            // Global
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Global {result.Key} As String";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for global {result.Key}");
                }
            }

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Public {result.Key} As String";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Private {result.Key} As String";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsVariableInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var declarationCode = $"Public {result.Key} As String";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, $"{declarationCode}\n\n{_moduleCode}").Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLocalVariableInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Public Sub Qux()
    Dim {result.Key} as String
End Sub";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for global {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLocalVariableInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Public Sub Qux()
    Dim {result.Key} as String
End Sub";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for global {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLocalVariableInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var code =
                    $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
Dim {LocalVariableName} as String
Const {LocalConstantName} as String = """"
{LineLabelName}:
End Sub

Public Function {FunctionName}()
Dim {result.Key} As String
End Function

Public Property Get {PropertyGetName}()
End Property

Public Property Let {PropertySetName}(v As Variant)
End Property

Public Property Set {PropertyLetName}(s As String)
End Property";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, code).Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsConstantInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Global
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Global Const {result.Key} As String = \"\"";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for global {result.Key}");
                }
            }

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Public Const {result.Key} As String= \"\"";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Private Const {result.Key} As String= \"\"";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(0, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsConstantInUserProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            // Global
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Global Const {result.Key} As String = \"\"";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for global {result.Key}");
                }
            }

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Public Const {result.Key} As String= \"\"";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Private Const {result.Key} As String= \"\"";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsConstantInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var declarationCode = $"Public Const {result.Key} As String= \"\"";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, $"{declarationCode}\n\n{_moduleCode}").Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLocalConstantInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Public Sub Qux()
    Const {result.Key} as String = """"
End Sub";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for global {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLocalConstantInUserProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Public Sub Qux()
    Const {result.Key} as String = """"
End Sub";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for global {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLocalConstantInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var code =
                    $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
Dim {LocalVariableName} as String
Const {LocalConstantName} as String = """"
{LineLabelName}:
End Sub

Public Function {FunctionName}()
Const {result.Key} as String = """"
End Function

Public Property Get {PropertyGetName}()
End Property

Public Property Let {PropertySetName}(v As Variant)
End Property

Public Property Set {PropertyLetName}(s As String)
End Property";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, code).Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Public Enum {result.Key}
    i
End Enum";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Private Enum {result.Key}
    i
End Enum";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(0, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationInUserProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Public Enum {result.Key}
    i
End Enum";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Private Enum {result.Key}
    i
End Enum";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var declarationCode =
                    $@"Public Enum {result.Key}
    i
End Enum";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, $"{declarationCode}\n\n{_moduleCode}").Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationMemberInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Public Enum Baz
    {result.Key}
End Enum";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Private Enum Baz
    {result.Key}
End Enum";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(0, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationMemberInUserProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Public Enum Baz
    {result.Key}
End Enum";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Private Enum Baz
    {result.Key}
End Enum";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationMemberInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var declarationCode =
                    $@"Public Enum Baz
    {result.Key}
End Enum";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, $"{declarationCode}\n\n{_moduleCode}").Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsResult_EnumerationWithSameNameAsEnumerationMember()
        {
            var code =
                @"Public enum SameName
    Baz
End Enum

Public enum Qux
    SameName
End Enum";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.StandardModule, code).Build();

            var vbe = builder.AddProject(project).Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsResult_EnumerationMemberWithSameNameAsEnumeration()
        {
            var code =
                @"Public enum Baz
    SameName
End Enum

Public enum SameName
    Qux
End Enum";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.StandardModule, code).Build();

            var vbe = builder.AddProject(project).Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EnumerationMemberWithSameNameAsEnumerationMember()
        {
            var code =
                @"Public enum Baz
    SameName
End Enum

Public enum Qux
    SameName
End Enum";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.StandardModule, code).Build();

            var vbe = builder.AddProject(project).Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EnumerationWithSameNameAsOwnMember()
        {
            var code =
                @"Public enum SameName
    SameName
End Enum";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.StandardModule, code).Build();

            var vbe = builder.AddProject(project).Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserDefinedTypeInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 1,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Public Type {result.Key}
    s As String
End Type";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Private Type {result.Key}
    s As String
End Type";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(0, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserDefinedTypeInUserProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 2,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Public Type {result.Key}
    s As String
End Type";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 1,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Private Type {result.Key}
    s As String
End Type";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserDefinedTypeInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var code =
                    $@"Public Type {result.Key}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
Dim {LocalVariableName} as String
Const {LocalConstantName} as String = """"
{LineLabelName}:
End Sub

Public Function {FunctionName}()
End Function

Public Property Get {PropertyGetName}()
End Property

Public Property Let {PropertySetName}(v As Variant)
End Property

Public Property Set {PropertyLetName}(s As String)
End Property";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, code).Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserDefinedTypeMemberInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Public Type T
    {result.Key} As String
End Type";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserDefinedTypeMemberInUserProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Public Type T
    {result.Key} As String
End Type";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserDefinedTypeMemberInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var declarationCode =
                    $@"Public Type T
    {result.Key} As String
End Type";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, $"{declarationCode}\n\n{_moduleCode}").Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLibraryProcedureInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Public Declare PtrSafe Sub {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Private Declare PtrSafe Sub {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(0, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLibraryProcedureInUserProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Public Declare PtrSafe Sub {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Private Declare PtrSafe Sub {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLibraryProcedureInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var declarationCode = $"Public Declare PtrSafe Sub {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, $"{declarationCode}\n\n{_moduleCode}").Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLibraryFunctionInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 0,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 1,
                [VariableName] = 1,
                [LocalVariableName] = 1,
                [ConstantName] = 1,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Public Declare PtrSafe Function {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode = $"Private Declare PtrSafe Function {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(0, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLibraryFunctionInUserProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 2,
                [FunctionName] = 2,
                [PropertyGetName] = 2,
                [PropertySetName] = 2,
                [PropertyLetName] = 2,
                [ParameterName] = 1,
                [VariableName] = 2,
                [LocalVariableName] = 1,
                [ConstantName] = 2,
                [LocalConstantName] = 1,
                [EnumerationName] = 2,
                [EnumerationMemberName] = 2,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 2,
                [LibraryFunctionName] = 2,
                [LineLabelName] = 0
            };

            // Public
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Public Declare PtrSafe Function {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }

            expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 1,
                [ProceduralModuleName] = 1,
                [ClassModuleName] = 0,
                [UserFormName] = 1,
                [DocumentName] = 1,
                [ProcedureName] = 1,
                [FunctionName] = 1,
                [PropertyGetName] = 1,
                [PropertySetName] = 1,
                [PropertyLetName] = 1,
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 1,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 1,
                [LibraryFunctionName] = 1,
                [LineLabelName] = 0
            };

            // Private
            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode = $"Private Declare PtrSafe Function {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for private {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLibraryFunctionInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 1,
                [VariableName] = 0,
                [LocalVariableName] = 1,
                [ConstantName] = 0,
                [LocalConstantName] = 1,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var declarationCode = $"Public Declare PtrSafe Function {result.Key} Lib \"lib.dll\" ()";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, $"{declarationCode}\n\n{_moduleCode}").Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLineLabelInReferencedProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var referencedModuleCode =
                    $@"Public Sub Qux()
    {result.Key}:
End Sub";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLineLabelInUserProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                [ProceduralModuleName] = 0,
                [ClassModuleName] = 0,
                [UserFormName] = 0,
                [DocumentName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var userModuleCode =
                    $@"Public Sub Qux()
    {result.Key}:
End Sub";

                var builder = new MockVbeBuilder();
                var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLineLabelInSameComponent()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProceduralModuleName] = 0,
                [ProcedureName] = 0,
                [FunctionName] = 0,
                [PropertyGetName] = 0,
                [PropertySetName] = 0,
                [PropertyLetName] = 0,
                [ParameterName] = 0,
                [VariableName] = 0,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 0,
                [EnumerationMemberName] = 0,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            foreach (var result in expectedResultCountsByDeclarationIdentifierName)
            {
                var code =
                    $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

Public Enum {EnumerationName}
    {EnumerationMemberName}
End Enum

Public Sub {ProcedureName}({ParameterName} As String)
Dim {LocalVariableName} as String
Const {LocalConstantName} as String = """"
{LineLabelName}:
End Sub

Public Function {FunctionName}()
{result.Key}:
End Function

Public Property Get {PropertyGetName}()
End Property

Public Property Let {PropertySetName}(v As Variant)
End Property

Public Property Set {PropertyLetName}(s As String)
End Property";

                var builder = new MockVbeBuilder();
                var project = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(ProceduralModuleName, ComponentType.StandardModule, code).Build();

                var vbe = builder.AddProject(project).Build();
                using (var state = MockParser.CreateAndParse(vbe.Object))
                {

                    var inspection = new ShadowedDeclarationInspection(state);
                    var inspectionResults = inspection.GetInspectionResults();

                    Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong number of inspection results for public {result.Key}");
                }
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideOptionPrivateModuleInReferencedProject()
        {
            var referencedModuleCode = $"Option Private Module\n\n{_moduleCode}";

            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                // Module name matters, because it can be shadowed without 'Option Private Module' statement
                .AddComponent(ProceduralModuleName, ComponentType.StandardModule, referencedModuleCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsInsideOptionPrivateModuleInUserProject()
        {
            var userModuleCode = $"Option Private Module\n\n{_moduleCode}";

            var builder = new MockVbeBuilder();
            var userProject = CreateUserProject(builder).AddComponent("Foo", ComponentType.StandardModule, userModuleCode).Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(12, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideClassModuleInReferencedProject()
        {
            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.ClassModule, _classCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideClassModuleInUserProject()
        {
            var builder = new MockVbeBuilder();
            var userProject = CreateUserProject(builder).AddComponent(ProceduralModuleName, ComponentType.ClassModule, _classCode).Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideUserFormInReferencedProject()
        {
            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.UserForm, _classCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideUserFormInUserProject()
        {
            var builder = new MockVbeBuilder();
            var userProject = CreateUserProject(builder).AddComponent(ProceduralModuleName, ComponentType.UserForm, _classCode).Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideDocumentInReferencedProject()
        {
            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.Document, _classCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideDocumentInUserProject()
        {
            var builder = new MockVbeBuilder();
            var userProject = CreateUserProject(builder).AddComponent(ProceduralModuleName, ComponentType.Document, _classCode).Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EventParameterWithSameNameAsDeclarationInReferencedProject()
        {
            const string sameName = "SameName";

            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.StandardModule, $"Public {sameName} As String")
                .Build();
            builder.AddProject(referencedProject);
            var userProject = builder.ProjectBuilder("Baz", ProjectProtection.Unprotected)
                .AddComponent("Qux", ComponentType.ClassModule, $"Public Event E ({sameName} As String)")
                .AddReference("Foo", "")
                .Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EventParameterWithSameNameAsDeclarationInUserProject()
        {
            const string sameName = "SameName";

            var builder = new MockVbeBuilder();
            var userProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.StandardModule, $"Public {sameName} As String")
                .AddComponent("Baz", ComponentType.ClassModule, $"Public Event E ({sameName} As String)")
                .Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EventParameterWithSameNameAsDeclarationInSameComponent()
        {
            const string sameName = "SameName";

            var code =
                $@"Public Event E ({sameName} As String)
Public {sameName} As String";

            var builder = new MockVbeBuilder();
            var userProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Baz", ComponentType.ClassModule, code)
                .Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_Ignored_DoesNotReturnResult()
        {
            string ignoredDeclarationCode =
                $@"'@Ignore ShadowedDeclaration
Public Sub {ProcedureName}()
End Sub";

            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.StandardModule, _moduleCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = builder.ProjectBuilder("Baz", ProjectProtection.Unprotected)
                .AddComponent("Qux", ComponentType.StandardModule, ignoredDeclarationCode)
                .Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        private void AssertResultCountsEqualForThoseWithExpectation(Dictionary<string, int> expectedResultCounts,
            Dictionary<string, int> actualResultCounts)
        {
            foreach (var expectedResultCount in expectedResultCounts)
            {
                var expectedCount = expectedResultCount.Value;
                int actualCount;
                if (!actualResultCounts.TryGetValue(expectedResultCount.Key, out actualCount))
                {
                    actualCount = 0;
                }
                Assert.AreEqual(expectedCount, actualCount,
                    $"Wrong number of inspection results for {expectedResultCount.Key}");
            }
        }

        private Dictionary<string, int> InspectionResultCountsByTargetIdentifierName(IInspection inspection)
        {
            var inspectionResults = inspection.GetInspectionResults();
            return InspectionResultCountsByTargetIdentifierName(inspectionResults);
        }

        private Dictionary<string, int> InspectionResultCountsByTargetIdentifierName(IEnumerable<IInspectionResult> inspectionResults)
        {
            return inspectionResults.GroupBy(result => result.Target.IdentifierName)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private MockProjectBuilder CreateUserProject(MockVbeBuilder builder, string projectName = ProjectName)
        {
            return builder.ProjectBuilder(projectName, ProjectProtection.Unprotected)
                .AddComponent(ProceduralModuleName, ComponentType.StandardModule, _moduleCode)
                .AddComponent(ClassModuleName, ComponentType.ClassModule, $"Public Event {EventName}()")
                .AddComponent(UserFormName, ComponentType.UserForm, "")
                .AddComponent(DocumentName, ComponentType.Document, "");
        }

        private MockVbeBuilder TestVbeWithUserProjectAndReferencedProjectWithComponentsOfOneType(string referencedProjectName, IEnumerable<string> testBaseNames, ComponentType referencedComponentsComponentType, Func<string, string> componentNameSelector, Func<string, string> componentCodeSelector, string userProjectName = ProjectName)
        {
            var builder = new MockVbeBuilder();
            var referencedProjectBuilder = builder.ProjectBuilder(referencedProjectName, ProjectProtection.Unprotected);

            foreach (var baseName in testBaseNames)
            {
                referencedProjectBuilder.AddComponent(componentNameSelector(baseName), referencedComponentsComponentType, componentCodeSelector(baseName));
            }
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder, userProjectName).AddReference(referencedProjectName, string.Empty).Build();
            builder.AddProject(userProject);

            return builder;
        }

        private MockVbeBuilder TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(string referencedProjectName, string referencedComponentName, ComponentType referencedComponentsComponentType, IEnumerable<string> testBaseNames, Func<string, string> moduleBodyElementCodeSelector, string userProjectName = ProjectName)
        {
            var builder = new MockVbeBuilder();
            var referencedProjectBuilder = builder.ProjectBuilder(referencedProjectName, ProjectProtection.Unprotected);

            var componentCode =
                ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(string.Empty, testBaseNames,
                    moduleBodyElementCodeSelector);

            referencedProjectBuilder.AddComponent(referencedComponentName, referencedComponentsComponentType, componentCode);
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder, userProjectName).AddReference(referencedProjectName, string.Empty).Build();
            builder.AddProject(userProject);

            return builder;
        }

        private MockVbeBuilder TestVbeWithUserProjectWithAdditionalComponentsOfOneType(IEnumerable<string> testBaseNames, ComponentType additionalComponentsComponentType, Func<string, string> componentNameSelector, Func<string, string> componentCodeSelector, string userProjectName = ProjectName)
        {
            var builder = new MockVbeBuilder();
            var userProjectBuilder = CreateUserProject(builder, userProjectName);

            foreach (var baseName in testBaseNames)
            {
                userProjectBuilder.AddComponent(componentNameSelector(baseName), additionalComponentsComponentType, componentCodeSelector(baseName));
            }

            var userProject = userProjectBuilder.Build();
            builder.AddProject(userProject);

            return builder;
        }

        private MockVbeBuilder TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(string additionalComponentName, ComponentType additionalComponentsComponentType, IEnumerable<string> testBaseNames, Func<string, string> moduleBodyElementCodeSelector, string userProjectName = ProjectName)
        {
            var builder = new MockVbeBuilder();
            var userProjectBuilder = CreateUserProject(builder, userProjectName);

            var componentCode =
                ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(string.Empty, testBaseNames,
                    moduleBodyElementCodeSelector);

            userProjectBuilder.AddComponent(additionalComponentName, additionalComponentsComponentType, componentCode);
            var userProject = userProjectBuilder.Build();
            builder.AddProject(userProject);

            return builder;
        }

        private string ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(string baseCode,
            IEnumerable<string> testBaseNames, Func<string, string> moduleBodyElementCodeSelector)
        {
            var codeBuilder = new StringBuilder();

            if (!string.Equals(baseCode, string.Empty))
            {
                codeBuilder.AppendLine(baseCode);
                codeBuilder.AppendLine();
            }

            foreach (var baseName in testBaseNames)
            {
                codeBuilder.AppendLine(moduleBodyElementCodeSelector(baseName));
                codeBuilder.AppendLine();
            }

            return codeBuilder.ToString();
        }
    }
}
