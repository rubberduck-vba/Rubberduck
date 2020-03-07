using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ShadowedDeclarationInspectionTests : InspectionTestsBase
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

        [Test]
        [Category("Inspections")]
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
                var referencedProject = builder.ProjectBuilder(expectedResultCount.Key, "somePath", ProjectProtection.Unprotected)
                    .AddComponent("Foo" + expectedResultCount.Key, ComponentType.StandardModule, string.Empty)
                    .Build();
                builder.AddProject(referencedProject);
                userProjectBuilder = userProjectBuilder.AddReference(expectedResultCount.Key, "somePath", 0, 0);
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

        [Test]
        [Category("Inspections")]
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
                    inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                }

                Assert.AreEqual(expectedResultCount.Value, inspectionResults.Count(), $"Wrong number of inspection results for {expectedResultCount.Key}");
            }
        }

        [Test]
        [Category("Inspections")]
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

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsProceduralModuleInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                //We do not test the following, because they cannot exist in the VBE.
                //[ProceduralModuleName] = 0,
                //[ClassModuleName] = 0,
                //[UserFormName] = 0,
                //[DocumentName] = 0,
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

        [Test]
        [Category("Inspections")]
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

        [Test]
        [Category("Inspections")]
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

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserFormInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                //We do not test the following, because they cannot exist in the VBE.
                //[ProceduralModuleName] = 0,
                //[ClassModuleName] = 0,
                //[UserFormName] = 0,
                //[DocumentName] = 0,
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

        [Test]
        [Category("Inspections")]
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

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsDocumentInContainingProject()
        {
            var expectedResultCountsByDeclarationIdentifierName = new Dictionary<string, int>
            {
                [ProjectName] = 0,
                //We do not test the following, because they cannot exist in the VBE.
                //[ProceduralModuleName] = 0,
                //[ClassModuleName] = 0,
                //[UserFormName] = 0,
                //[DocumentName] = 0,
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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

Public Property Let {PropertyLetName}(v As Variant)
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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

        [Test]
        [Category("Inspections")]
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
                referencedComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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
                additionalComponentComponentType: ComponentType.StandardModule,
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

        [Test]
        [Category("Inspections")]
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

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsGlobalVariableInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Global {key} As String");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicVariableInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public {key} As String");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateVariableInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private {key} As String");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsGlobalVariableInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Global {key} As String");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicVariableInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public {key} As String");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateVariableInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private {key} As String");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var baseCode =
                $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

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
End Property

Public Property Let {PropertyLetName}(v As Variant)
End Property";

            var code = ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public {key} As String");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub Qux{key}()
    Dim {key} as String
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

        [Test]
        [Category("Inspections")]
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub Qux{key}()
    Dim {key} as String
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

        [Test]
        [Category("Inspections")]
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

            var code = ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(
                baseCode: _moduleCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Function Foo{key}()
    Dim {key} As String
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

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsGlobalConstantInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Global Const {key} As String = """"");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicConstantInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Const {key} As String = """"");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateConstantInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Const {key} As String = """"");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsGlobalConstantInUserProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Global Const {key} As String = """"");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicConstantInUserProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Const {key} As String = """"");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateConstantInUserProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Const {key} As String = """"");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var baseCode =
                $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public {VariableName} As String

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
End Property

Public Property Let {PropertyLetName}(v As Variant)
End Property";

            var code = ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Const {key} As String = """"");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub Qux{key}()
    Const {key} as String = """"
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

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLocalConstantInContainingProject()
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
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub Qux{key}()
    Const {key} as String = """"
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

        [Test]
        [Category("Inspections")]
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

            var code = ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(
                baseCode: _moduleCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Function Foo{key}()
    Const {key} As String = """"
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

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicEnumerationInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Enum {key}
    i{key}
End Enum");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateEnumerationInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Enum {key}
    i{key}
End Enum");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicEnumerationInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Enum {key}
    i{key}
End Enum");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateEnumerationInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Enum {key}
    i{key}
End Enum");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var baseCode =
                $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public Const {ConstantName} As String = """"

Public {VariableName} As String

Public Enum FooBar
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
End Property

Public Property Let {PropertyLetName}(v As Variant)
End Property";

            var code = ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Enum {key}
    i{key}
End Enum");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
        ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationMemberOfPublicEnumerationInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Enum Baz{key}
    {key}
End Enum");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationMemberOfPrivateEnumerationInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Enum Baz{key}
    {key}
End Enum");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationMemberOrPublicEnumerationInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Enum Baz{key}
    {key}
End Enum");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationMemberOrPrivateEnumerationInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Enum Baz{key}
    {key}
End Enum");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var code = ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(
                baseCode: _moduleCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Enum Baz{key}
    {key}
End Enum");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsResult_EnumerationWithSameNameAsEnumerationMember()
        {
            var code =
                @"Public Enum SameName
    Baz
End Enum

Public Enum Qux
    SameName
End Enum";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsResult_EnumerationMemberWithSameNameAsEnumeration()
        {
            var code =
                @"Public Enum Baz
    SameName
End Enum

Public Enum SameName
    Qux
End Enum";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EnumerationMemberWithSameNameAsEnumerationMember()
        {
            var code =
                @"Public Enum Baz
    SameName
End Enum

Public Enum Qux
    SameName
End Enum";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EnumerationWithSameNameAsOwnMember()
        {
            var code =
                @"Public enum SameName
    SameName
End Enum";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicUserDefinedTypeInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Type {key}
    s{key} As String
End Type");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateUserDefinedTypeInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Type {key}
    s{key} As String
End Type");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicUserDefinedTypeInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Type {key}
    s{key} As String
End Type");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateUserDefinedTypeInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Type {key}
    s{key} As String
End Type");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var baseCode = $@"Public Type FooBar
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

            var code = ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Type {key}
    s{key} As String
End Type");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Type T{key}
    {key} As String
End Type");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserDefinedTypeMemberInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Type T{key}
    {key} As String
End Type");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var code = ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(
                baseCode: _moduleCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Type T{key}
    {key} As String
End Type");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_UserDefinedTypeWithSameNameAsOwnMember()
        {
            var code =
                @"Public Type SameName
    SameName As String
End Type";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicLibraryProcedureInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Declare PtrSafe Sub {key} Lib ""lib.dll"" ()");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateLibraryProcedureInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Declare PtrSafe Sub {key} Lib ""lib.dll"" ()");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicLibraryProcedureInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Declare PtrSafe Sub {key} Lib ""lib.dll"" ()");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateLibraryProcedureInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Declare PtrSafe Sub {key} Lib ""lib.dll"" ()");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var baseCode =
                $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll"" ()

Public Const {ConstantName} As String = """"

Public {VariableName} As String

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
End Property

Public Property Let {PropertyLetName}(v As Variant)
End Property";

            var code = ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Declare PtrSafe Sub {key} Lib ""lib.dll"" ()");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicLibraryFunctionInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Declare PtrSafe Function {key} Lib ""lib.dll"" ()");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateLibraryFunctionInReferencedProject()
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Declare PtrSafe Function {key} Lib ""lib.dll"" ()");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void
            ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPublicLibraryFunctionInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Declare PtrSafe Function {key} Lib ""lib.dll"" ()");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPrivateLibraryFunctionInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Private Declare PtrSafe Function {key} Lib ""lib.dll"" ()");
            var vbe = builder.Build();

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var baseCode =
                $@"Public Type {UserDefinedTypeName}
    {UserDefinedTypeMemberName} As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Const {ConstantName} As String = """"

Public {VariableName} As String

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
End Property

Public Property Let {PropertyLetName}(v As Variant)
End Property";

            var code = ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(
                baseCode: baseCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleDeclarationCodeSelector: key => $@"Public Declare PtrSafe Function {key} Lib ""lib.dll"" ()");
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _);

            Dictionary<string, int> inspectionResultCounts;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspection);
            }

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
        }

        [Test]
        [Category("Inspections")]
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

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub Qux{key}()
{key}:
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

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLineLabelInContainingProject()
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

            var builder = TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Sub Qux{key}()
{key}:
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

        [Test]
        [Category("Inspections")]
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

            var code = ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(
                baseCode: _moduleCode,
                testBaseNames: expectedResultCountsByDeclarationIdentifierName.Keys,
                moduleBodyElementCodeSelector: key => $@"Public Function Foo{key}()
{key}:
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

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideOptionPrivateModuleInReferencedProject()
        {
            var referencedComponentCodeBuilder = new StringBuilder();
            referencedComponentCodeBuilder.AppendLine("Option Private Module");
            referencedComponentCodeBuilder.AppendLine();
            referencedComponentCodeBuilder.Append(_moduleCode);

            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponent(
                referencedProjectName: "Foo",
                referencedComponentName: ProceduralModuleName, // Module name matters, because it can be shadowed without 'Option Private Module' statement
                referencedComponentComponentType: ComponentType.StandardModule,
                referencedComponentCode: referencedComponentCodeBuilder.ToString());
            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsInsideOptionPrivateModuleInContainingProject()
        {
            var additionalComponentCodeBuilder = new StringBuilder();
            additionalComponentCodeBuilder.AppendLine("Option Private Module");
            additionalComponentCodeBuilder.AppendLine();
            additionalComponentCodeBuilder.Append(_moduleCode);

            var builder = TestVbeWithUserProjectWithAdditionalComponent(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.StandardModule,
                additionalComponentCode: additionalComponentCodeBuilder.ToString());
            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.AreEqual(12, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideClassModuleInReferencedProject()
        {
            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponent(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.ClassModule,
                referencedComponentCode: _classCode);
            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsInsideClassModuleInContainingProject()
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
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };
            
            var builder = TestVbeWithUserProjectWithAdditionalComponent(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.ClassModule,
                additionalComponentCode: _classCode);
            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None).ToList();
            }
            var inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspectionResults);

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
            //All shadowing happens inside the class module.
            Assert.IsTrue(inspectionResults.All(result => result.Target.QualifiedName.QualifiedModuleName.ComponentName == "Foo"));
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideUserFormInReferencedProject()
        {
            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponent(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.UserForm,
                referencedComponentCode: _classCode);
            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsInsideUserFormInContainingProject()
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
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponent(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.UserForm,
                additionalComponentCode: _classCode);
            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None).ToList();
            }
            var inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspectionResults);

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
            //All shadowing happens inside the user form.
            Assert.IsTrue(inspectionResults.All(result => result.Target.QualifiedName.QualifiedModuleName.ComponentName == "Foo"));
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideDocumentInReferencedProject()
        {
            var builder = TestVbeWithUserProjectAndReferencedProjectWithComponent(
                referencedProjectName: "Foo",
                referencedComponentName: "Bar",
                referencedComponentComponentType: ComponentType.Document,
                referencedComponentCode: _classCode);
            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsInsideDocumentInContainingProject()
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
                [ParameterName] = 0,
                [VariableName] = 1,
                [LocalVariableName] = 0,
                [ConstantName] = 0,
                [LocalConstantName] = 0,
                [EnumerationName] = 1,
                [EnumerationMemberName] = 1,
                [UserDefinedTypeName] = 0,
                [UserDefinedTypeMemberName] = 0,
                [LibraryProcedureName] = 0,
                [LibraryFunctionName] = 0,
                [LineLabelName] = 0
            };

            var builder = TestVbeWithUserProjectWithAdditionalComponent(
                additionalComponentName: "Foo",
                additionalComponentComponentType: ComponentType.Document,
                additionalComponentCode: _classCode);
            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None).ToList();
            }
            var inspectionResultCounts = InspectionResultCountsByTargetIdentifierName(inspectionResults);

            AssertResultCountsEqualForThoseWithExpectation(expectedResultCountsByDeclarationIdentifierName, inspectionResultCounts);
            //All shadowing happens inside the document module.
            Assert.IsTrue(inspectionResults.All(result => result.Target.QualifiedName.QualifiedModuleName.ComponentName == "Foo"));
        }

        [Test]
        [Category("Inspections")]
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
                .AddReference("Foo", string.Empty, 0, 0)
                .Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EventParameterWithSameNameAsDeclarationInContainingProject()
        {
            const string sameName = "SameName";

            var builder = new MockVbeBuilder();
            var userProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.StandardModule, $"Public {sameName} As String")
                .AddComponent("Baz", ComponentType.ClassModule, $"Public Event E ({sameName} As String)")
                .Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EventParameterWithSameNameAsDeclarationInSameComponent()
        {
            const string sameName = "SameName";

            var code =
                $@"Public Event E ({sameName} As String)
Public {sameName} As String";

            var vbe = MockVbeBuilder.BuildFromSingleModule(code, ComponentType.ClassModule, out _);

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("", "As Object", "")]    //Variable
        [TestCase("Function", "As Object", "End Function")]
        [TestCase("Sub", "", "End Sub")]
        [TestCase("Property Get", "As Object", "End Property")]
        [TestCase("Property Let", "", "End Property")]
        [TestCase("Property Set", "", "End Property")]
        public void ShadowedDeclaration_DoesNotReturnResult_AssertBecauseOfDebugAssert(string memberType, string asTypeClause, string endTag)
        {
            var code =
                $@"Public {memberType} Assert() {asTypeClause} 
{endTag}";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestClass", ComponentType.ClassModule, code)
                .AddReference(ReferenceLibrary.VBA)
                .AddProjectToVbeBuilder()
                .Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_LocalAssertVariableAssertBecauseOfDebugAssert()
        {
            var code =
                @"Public Sub Foo()
    Dim Assert As Long
End Sub";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestClass", ComponentType.ClassModule, code)
                .AddReference(ReferenceLibrary.VBA)
                .AddProjectToVbeBuilder()
                .Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_AssertParameterAssertBecauseOfDebugAssert()
        {
            var code =
                @"Public Sub Foo(Assert As Boolean)
End Sub";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestClass", ComponentType.ClassModule, code)
                .AddReference(ReferenceLibrary.VBA)
                .AddProjectToVbeBuilder()
                .Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("", "As Object", "")]    //Variable
        [TestCase("Function", "As Object", "End Function")]
        [TestCase("Sub", "", "End Sub")]
        [TestCase("Property Get", "As Object", "End Property")]
        [TestCase("Property Let", "", "End Property")]
        [TestCase("Property Set", "", "End Property")]
        public void ShadowedDeclaration_ReturnResult_LocalAssertVariableAssertBecauseOfNonDebugAssert(string memberType, string asTypeClause, string endTag)
        {
            var code =
                $@"Public {memberType} Assert() {asTypeClause} 
{endTag}

Public Sub Foo()
    Dim Assert As Long
End Sub";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestClass", ComponentType.ClassModule, code)
                .AddReference(ReferenceLibrary.VBA)
                .AddProjectToVbeBuilder()
                .Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.AreEqual(1,inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("", "As Object", "")]    //Variable
        [TestCase("Function", "As Object", "End Function")]
        [TestCase("Sub", "", "End Sub")]
        [TestCase("Property Get", "As Object", "End Property")]
        [TestCase("Property Let", "", "End Property")]
        [TestCase("Property Set", "", "End Property")]
        public void ShadowedDeclaration_ReturnResult_AssertParameterAssertBecauseOfNonDebugAssert(string memberType, string asTypeClause, string endTag)
        {
            var code =
                $@"Public {memberType} Assert() {asTypeClause} 
{endTag}

Public Sub Foo(Assert As Boolean)
End Sub";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestClass", ComponentType.ClassModule, code)
                .AddReference(ReferenceLibrary.VBA)
                .AddProjectToVbeBuilder()
                .Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
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
                .AddReference("Foo", string.Empty, 0, 0)
                .Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();

            IEnumerable<IInspectionResult> inspectionResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ShadowedDeclarationInspection(state);
                inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ShadowedDeclarationInspection(null);

            Assert.AreEqual("ShadowedDeclarationInspection", inspection.Name);
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
            var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
            return InspectionResultCountsByTargetIdentifierName(inspectionResults);
        }

        private Dictionary<string, int> InspectionResultCountsByTargetIdentifierName(IEnumerable<IInspectionResult> inspectionResults)
        {
            return inspectionResults.GroupBy(result => result.Target.IdentifierName)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private MockProjectBuilder CreateUserProject(MockVbeBuilder builder, string projectName = ProjectName, string projectPath = "")
        {
            return builder.ProjectBuilder(projectName, projectPath, ProjectProtection.Unprotected)
                .AddComponent(ProceduralModuleName, ComponentType.StandardModule, _moduleCode)
                .AddComponent(ClassModuleName, ComponentType.ClassModule, $"Public Event {EventName}()")
                .AddComponent(UserFormName, ComponentType.UserForm, "")
                .AddComponent(DocumentName, ComponentType.Document, "");
        }

        private MockVbeBuilder TestVbeWithUserProjectAndReferencedProjectWithComponentsOfOneType(string referencedProjectName, IEnumerable<string> testBaseNames, ComponentType referencedComponentsComponentType, Func<string, string> componentNameSelector, Func<string, string> componentCodeSelector, string userProjectName = ProjectName)
        {
            var builder = new MockVbeBuilder();
            var project = string.IsNullOrEmpty(referencedProjectName) ? "Irrelevant" : referencedProjectName;
            var path = $@"C:\{project}.xlsm";
            var referencedProjectBuilder = builder.ProjectBuilder(referencedProjectName, path, ProjectProtection.Unprotected);

            foreach (var baseName in testBaseNames)
            {
                referencedProjectBuilder.AddComponent(componentNameSelector(baseName), referencedComponentsComponentType, componentCodeSelector(baseName));
            }

            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var userProject = CreateUserProject(builder, userProjectName).AddReference(project, path, 0, 0).Build();
            builder.AddProject(userProject);

            return builder;
        }

        private MockVbeBuilder TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleBodyElements(string referencedProjectName, string referencedComponentName, ComponentType referencedComponentComponentType, IEnumerable<string> testBaseNames, Func<string, string> moduleBodyElementCodeSelector, string userProjectName = ProjectName)
        {
            var componentCode =
                ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(string.Empty, testBaseNames,
                    moduleBodyElementCodeSelector);

            return TestVbeWithUserProjectAndReferencedProjectWithComponent(referencedProjectName,referencedComponentName,referencedComponentComponentType, componentCode, userProjectName);
        }

        private MockVbeBuilder TestVbeWithUserProjectAndReferencedProjectWithComponentWithSelectedModuleDeclarations(string referencedProjectName, string referencedComponentName, ComponentType referencedComponentComponentType, IEnumerable<string> testBaseNames, Func<string, string> moduleDeclarationCodeSelector, string userProjectName = ProjectName)
        {
            var componentCode =
                ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(string.Empty, testBaseNames,
                    moduleDeclarationCodeSelector);

            return TestVbeWithUserProjectAndReferencedProjectWithComponent(referencedProjectName, referencedComponentName, referencedComponentComponentType, componentCode, userProjectName);
        }

        private MockVbeBuilder TestVbeWithUserProjectAndReferencedProjectWithComponent(string referencedProjectName,
            string referencedComponentName, 
            ComponentType referencedComponentComponentType,
            string referencedComponentCode,
            string userProjectName = ProjectName)
        {
            var project = string.IsNullOrEmpty(referencedProjectName) ? "Irrelevant" : referencedProjectName;
            var path = $@"C:\{project}.xlsm";

            var builder = new MockVbeBuilder();
            var referencedProjectBuilder = builder.ProjectBuilder(referencedProjectName, path, ProjectProtection.Unprotected);
            referencedProjectBuilder.AddComponent(referencedComponentName, referencedComponentComponentType, referencedComponentCode);
            var referencedProject = referencedProjectBuilder.Build();
            builder.AddProject(referencedProject);

            var userProject = CreateUserProject(builder, userProjectName).AddReference(project, path, 0, 0).Build();
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

        private MockVbeBuilder TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleBodyElements(string additionalComponentName, ComponentType additionalComponentComponentType, IEnumerable<string> testBaseNames, Func<string, string> moduleBodyElementCodeSelector, string userProjectName = ProjectName)
        {
            var componentCode =
                ModuleCodeFromBaseCodeAndSelectedModuleBodyElements(string.Empty, testBaseNames,
                    moduleBodyElementCodeSelector);

            return TestVbeWithUserProjectWithAdditionalComponent(additionalComponentName, additionalComponentComponentType, componentCode, userProjectName);
        }

        private MockVbeBuilder TestVbeWithUserProjectWithAdditionalComponentWithSelectedModuleDeclarations(string additionalComponentName, ComponentType additionalComponentComponentType, IEnumerable<string> testBaseNames, Func<string, string> moduleDeclarationCodeSelector, string userProjectName = ProjectName)
        {
            var componentCode =
                ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(string.Empty, testBaseNames,
                    moduleDeclarationCodeSelector);

            return TestVbeWithUserProjectWithAdditionalComponent(additionalComponentName, additionalComponentComponentType, componentCode, userProjectName);
        }

        private MockVbeBuilder TestVbeWithUserProjectWithAdditionalComponent(string additionalComponentName, ComponentType additionalComponentComponentType, string additionalComponentCode, string userProjectName = ProjectName)
        {
            var builder = new MockVbeBuilder();
            var userProjectBuilder = CreateUserProject(builder, userProjectName);
            userProjectBuilder.AddComponent(additionalComponentName, additionalComponentComponentType, additionalComponentCode);
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

        private string ModuleCodeFromSelectedModuleDeclarationsAndBaseCode(string baseCode,
            IEnumerable<string> testBaseNames, Func<string, string> moduleDeclarationCodeSelector)
        {
            var codeBuilder = new StringBuilder();

            foreach (var baseName in testBaseNames)
            {
                codeBuilder.AppendLine(moduleDeclarationCodeSelector(baseName));
                codeBuilder.AppendLine();
            }

            if (!string.Equals(baseCode, string.Empty))
            {
                codeBuilder.Append(baseCode);
            }

            return codeBuilder.ToString();
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ShadowedDeclarationInspection(state);
        }
    }
}
