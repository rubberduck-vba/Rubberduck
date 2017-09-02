using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
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
        private const string ConstantName = "SameNameConstant";
        private const string EnumerationName = "SameNameEnumeration";
        private const string EnumerationMemberName = "SameNameEnumerationMember";
        private const string EventName = "SameNameEvent";
        private const string UserDefinedTypeName = "SameNameUserDefinedType";
        private const string LibraryProcedureName = "SameNameLibraryProcedure";
        private const string LibraryFunctionName = "SameNameLibraryFunction";
        private const string LineLabelName = "SameNameLineLabel";

        private readonly string moduleCode = 
$@"Public Type {UserDefinedTypeName}
    s As String
End Type

Public Declare PtrSafe Sub {LibraryProcedureName} Lib ""lib.dll"" ()

Public Declare PtrSafe Function {LibraryFunctionName} Lib ""lib.dll""()

Public {VariableName} As String

Public Const {ConstantName} As String = """"

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

        private readonly string classCode =
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
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 0, [ProceduralModuleName] = 0, [ClassModuleName] = 0, [UserFormName] = 0, [DocumentName] = 0, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 0, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            foreach (var result in declarationResults)
            {
                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder(result.Key, ProjectProtection.Unprotected)
                    .AddComponent("Foo", ComponentType.StandardModule, "")
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference(result.Key, "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsProceduralModuleInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            foreach (var result in declarationResults)
            {
                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(result.Key, ComponentType.StandardModule, "")
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsNonExposedClassModuleInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 0, [ProceduralModuleName] = 0, [ClassModuleName] = 0, [UserFormName] = 0, [DocumentName] = 0, [ProcedureName] = 0, [FunctionName] = 0,
                [PropertyGetName] = 0, [PropertySetName] = 0, [PropertyLetName] = 0, [ParameterName] = 0, [VariableName] = 0, [ConstantName] = 0,
                [EnumerationName] = 0, [EnumerationMemberName] = 0, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 0, [LibraryFunctionName] = 0, [LineLabelName] = 0
            };

            foreach (var result in declarationResults)
            {
                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(result.Key, ComponentType.ClassModule, "")
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserFormInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 0, [ProceduralModuleName] = 0, [ClassModuleName] = 0, [UserFormName] = 0, [DocumentName] = 0, [ProcedureName] = 0, [FunctionName] = 0,
                [PropertyGetName] = 0, [PropertySetName] = 0, [PropertyLetName] = 0, [ParameterName] = 0, [VariableName] = 0, [ConstantName] = 0,
                [EnumerationName] = 0, [EnumerationMemberName] = 0, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 0, [LibraryFunctionName] = 0, [LineLabelName] = 0
            };

            foreach (var result in declarationResults)
            {
                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(result.Key, ComponentType.UserForm, "")
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsDocumentInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 0, [ProceduralModuleName] = 0, [ClassModuleName] = 0, [UserFormName] = 0, [DocumentName] = 0, [ProcedureName] = 0, [FunctionName] = 0,
                [PropertyGetName] = 0, [PropertySetName] = 0, [PropertyLetName] = 0, [ParameterName] = 0, [VariableName] = 0, [ConstantName] = 0,
                [EnumerationName] = 0, [EnumerationMemberName] = 0, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 0, [LibraryFunctionName] = 0, [LineLabelName] = 0
            };

            foreach (var result in declarationResults)
            {
                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent(result.Key, ComponentType.Document, "")
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsProcedureInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Public Sub {result.Key}()
End Sub";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }
            
            // Private
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Private Sub {result.Key}()
End Sub";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsFunctionInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Public Function {result.Key}()
End Function";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Private Function {result.Key}()
End Function";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPropertyGetInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Public Property Get {result.Key}() As String
End Property";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Private Property Get {result.Key}() As String
End Property";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPropertySetInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Public Property Set {result.Key}(v As Variant)
End Property";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Private Property Set {result.Key}(v As Variant)
End Property";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsPropertyLetInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Public Property Let {result.Key}(s As String)
End Property";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Private Property Let {result.Key}(s As String)
End Property";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsParameterInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 0, [ProceduralModuleName] = 0, [ClassModuleName] = 0, [UserFormName] = 0, [DocumentName] = 0, [ProcedureName] = 0, [FunctionName] = 0,
                [PropertyGetName] = 0, [PropertySetName] = 0, [PropertyLetName] = 0, [ParameterName] = 0, [VariableName] = 0, [ConstantName] = 0,
                [EnumerationName] = 0, [EnumerationMemberName] = 0, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 0, [LibraryFunctionName] = 0, [LineLabelName] = 0
            };

            foreach (var result in declarationResults)
            {
                var referencedModuleCode =
$@"Public Sub Qux({result.Key} As String)
End Sub";

                var builder = new MockVbeBuilder();
                var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                    .AddComponent("Bar", ComponentType.StandardModule, referencedModuleCode)
                    .Build();
                builder.AddProject(referencedProject);
                var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
                builder.AddProject(userProject);

                var vbe = builder.Build();
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsVariableInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Global
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for global {result.Key}");
            }

            // Public
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsConstantInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Global
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for global {result.Key}");
            }

            // Public
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsEnumerationMemberInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserDefinedTypeInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 0, [FunctionName] = 0,
                [PropertyGetName] = 0, [PropertySetName] = 0, [PropertyLetName] = 0, [ParameterName] = 0, [VariableName] = 0, [ConstantName] = 0,
                [EnumerationName] = 0, [EnumerationMemberName] = 0, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 0, [LibraryFunctionName] = 0, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsUserDefinedTypeMemberInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 0, [ProceduralModuleName] = 0, [ClassModuleName] = 0, [UserFormName] = 0, [DocumentName] = 0, [ProcedureName] = 0, [FunctionName] = 0,
                [PropertyGetName] = 0, [PropertySetName] = 0, [PropertyLetName] = 0, [ParameterName] = 0, [VariableName] = 0, [ConstantName] = 0,
                [EnumerationName] = 0, [EnumerationMemberName] = 0, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 0, [LibraryFunctionName] = 0, [LineLabelName] = 0
            };

            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLibraryProcedureInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLibraryFunctionInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 1, [ProceduralModuleName] = 1, [ClassModuleName] = 0, [UserFormName] = 1, [DocumentName] = 1, [ProcedureName] = 1, [FunctionName] = 1,
                [PropertyGetName] = 1, [PropertySetName] = 1, [PropertyLetName] = 1, [ParameterName] = 1, [VariableName] = 1, [ConstantName] = 1,
                [EnumerationName] = 1, [EnumerationMemberName] = 1, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 1, [LibraryFunctionName] = 1, [LineLabelName] = 0
            };

            // Public
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for public {result.Key}");
            }

            // Private
            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count(), $"Wrong inspection result for private {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_ReturnsCorrectResult_DeclarationsWithSameNameAsLineLabelInReferencedProject()
        {
            var declarationResults = new Dictionary<string, int>
            {
                [ProjectName] = 0, [ProceduralModuleName] = 0, [ClassModuleName] = 0, [UserFormName] = 0, [DocumentName] = 0, [ProcedureName] = 0, [FunctionName] = 0,
                [PropertyGetName] = 0, [PropertySetName] = 0, [PropertyLetName] = 0, [ParameterName] = 0, [VariableName] = 0, [ConstantName] = 0,
                [EnumerationName] = 0, [EnumerationMemberName] = 0, [UserDefinedTypeName] = 0, [LibraryProcedureName] = 0, [LibraryFunctionName] = 0, [LineLabelName] = 0
            };

            foreach (var result in declarationResults)
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
                var state = MockParser.CreateAndParse(vbe.Object);

                var inspection = new ShadowedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(result.Value, inspectionResults.Count(), $"Wrong inspection result for {result.Key}");
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideOptionPrivateModule()
        {
            var referencedModuleCode = $"Option Private Module\n\n{moduleCode}";

            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                // Module name matters, because it can be shadowed without 'Option Private Module' statement
                .AddComponent(ProceduralModuleName, ComponentType.StandardModule, referencedModuleCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ShadowedDeclarationInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideClassModule()
        {
            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.ClassModule, classCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ShadowedDeclarationInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideUserForm()
        {
            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.UserForm, classCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ShadowedDeclarationInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_DeclarationsInsideDocument()
        {
            var builder = new MockVbeBuilder();
            var referencedProject = builder.ProjectBuilder("Foo", ProjectProtection.Unprotected)
                .AddComponent("Bar", ComponentType.Document, classCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = CreateUserProject(builder).AddReference("Foo", "").Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ShadowedDeclarationInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ShadowedDeclaration_DoesNotReturnResult_EventParameterInUserCode()
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
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ShadowedDeclarationInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
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
                .AddComponent("Bar", ComponentType.StandardModule, moduleCode)
                .Build();
            builder.AddProject(referencedProject);
            var userProject = builder.ProjectBuilder("Baz", ProjectProtection.Unprotected)
                .AddComponent("Qux", ComponentType.StandardModule, ignoredDeclarationCode)
                .Build();
            builder.AddProject(userProject);

            var vbe = builder.Build();
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ShadowedDeclarationInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        private MockProjectBuilder CreateUserProject(MockVbeBuilder builder)
        {
            return builder.ProjectBuilder(ProjectName, ProjectProtection.Unprotected)
                .AddComponent(ProceduralModuleName, ComponentType.StandardModule, moduleCode)
                .AddComponent(ClassModuleName, ComponentType.ClassModule, $"Public Event {EventName}()")
                .AddComponent(UserFormName, ComponentType.UserForm, "")
                .AddComponent(DocumentName, ComponentType.Document, "");
        }
    }
}
