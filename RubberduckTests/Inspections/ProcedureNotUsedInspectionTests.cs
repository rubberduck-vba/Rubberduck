using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ProcedureNotUsedInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        [TestCase("Sub", "")]
        [TestCase("Function", "")]
        [TestCase("Property", " Get")]
        [TestCase("Property", " Let")]
        [TestCase("Property", " Set")]
        public void ProcedureNotUsed_ReturnsResult(string memberType, string memberTypeExtension)
        {
            var inputCode =
                $@"Private {memberType}{memberTypeExtension} Foo(arg As Variant)
End {memberType}";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureNotUsed_ReturnsResult_MultipleSubs()
        {
            const string inputCode =
                @"Private Sub Foo()
End Sub

Private Sub Goo()
End Sub";

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureUsed_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private Sub Foo()
    Goo
End Sub

Private Sub Goo()
    Foo
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureNotUsed_ReturnsResult_SomeProceduresUsed()
        {
            const string inputCode =
                @"Private Sub Foo()
End Sub

Private Sub Goo()
    Foo
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureNotUsed_DoesNotReturnResult_InterfaceImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";

            var modules = new (string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
            };

            var relevantResults = InspectionResultsForModules(modules)
                .Where(result => result.Target.DeclarationType == DeclarationType.Procedure
                    && result.Target.QualifiedModuleName.ComponentName == "Class1");

            Assert.AreEqual(0, relevantResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureNotUsed_ReturnsResult_ImplementedInterfaceMethod()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";

            var modules = new (string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
            };

            var results = InspectionResultsForModules(modules);
            var relevantResults = results
                .Where(result => result.Target.DeclarationType == DeclarationType.Procedure
                    && result.Target.QualifiedModuleName.ComponentName == "IClass1");

            Assert.AreEqual(1, relevantResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureNotUsed_ReturnsResult_MarkedInterfaceMethod()
        {
            const string inputCode1 =
                @"
'@Interface

Public Sub DoSomething(ByVal a As Integer)
End Sub
";


            var modules = new (string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule)
            };

            var results = InspectionResultsForModules(modules);
            var relevantResults = results
                .Where(result => result.Target.DeclarationType == DeclarationType.Procedure
                    && result.Target.QualifiedModuleName.ComponentName == "IClass1");

            Assert.AreEqual(1, relevantResults.Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Sub", "")]
        [TestCase("Function", "")]
        [TestCase("Property", " Get")]
        [TestCase("Property", " Let")]
        [TestCase("Property", " Set")]
        public void ProcedureNotUsed_DoesNotReturnResult_ExposedClass_Public(string memberType, string memberTypeExtension)
        {
            var inputCode =
                $@"
Attribute VB_Exposed = True

Public {memberType}{memberTypeExtension} Foo(arg As Variant)
End {memberType}";

            Assert.AreEqual(0, InspectionResultsForModules(("Class1", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Sub", "")]
        [TestCase("Function", "")]
        [TestCase("Property", " Get")]
        [TestCase("Property", " Let")]
        [TestCase("Property", " Set")]
        public void ProcedureNotUsed_ReturnsResult_ExposedClass_Private(string memberType, string memberTypeExtension)
        {
            var inputCode =
                $@"
Attribute VB_Exposed = True

Private {memberType}{memberTypeExtension} Foo(arg As Variant)
End {memberType}";

            Assert.AreEqual(1, InspectionResultsForModules(("Class1", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Sub", "")]
        [TestCase("Function", "")]
        [TestCase("Property", " Get")]
        [TestCase("Property", " Let")]
        [TestCase("Property", " Set")]
        public void ProcedureNotUsed_ReturnsResult_ExposedClass_Friend(string memberType, string memberTypeExtension)
        {
            var inputCode =
                $@"
Attribute VB_Exposed = True

Friend {memberType}{memberTypeExtension} Foo(arg As Variant)
End {memberType}";

            Assert.AreEqual(1, InspectionResultsForModules(("Class1", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureNotUsed_HandlerIsIgnoredForUnraisedEvent()
        {
            const string inputCode1 = @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";
            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            var modules = new(string, string, ComponentType)[] 
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count(result => result.Target.DeclarationType == DeclarationType.Procedure));
        }

        [TestCase("@TestMethod(\"TestCategory\")")]
        [TestCase("@ModuleInitialize")]
        [TestCase("@ModuleCleanup")]
        [TestCase("@TestInitialize")]
        [TestCase("@TestCleanup")]
        [Category("Inspections")]
        public void ProcedureNotUsed_NoResultForTestRelatedMethods(string annotationText)
        {
            string inputCode =
                $@"
'{annotationText}
Private Sub TestRelatedMethod()
End Sub";
            
            Assert.AreEqual(0, InspectionResultsForModules(("TestClass", inputCode, ComponentType.StandardModule)).Count());
        }

        [TestCase("Class_Initialize")]
        [TestCase("class_initialize")]
        [TestCase("Class_Terminate")]
        [TestCase("class_terminate")]
        [Category("Inspections")]
        public void ProcedureNotUsed_NoResultForLifeCycleHandlers(string subName)
        {
                string inputCode =
$@"Private Sub {subName}()
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("TestClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [TestCase(ComponentType.StandardModule, "auto_open", "module1", "Excel")]
        [TestCase(ComponentType.StandardModule, "auto_close", "module1", "Excel")]
        [TestCase(ComponentType.StandardModule, "AutoExec", "module1", "Word")]
        [TestCase(ComponentType.StandardModule, "AutoNew", "module1", "Word")]
        [TestCase(ComponentType.StandardModule, "AutoOpen", "module1", "Word")]
        [TestCase(ComponentType.StandardModule, "AutoClose", "module1", "Word")]
        [TestCase(ComponentType.StandardModule, "AutoExit", "module1", "Word")]
        [TestCase(ComponentType.Document, "AutoExec", "module1", "Word")]
        [TestCase(ComponentType.Document, "AutoNew", "module1", "Word")]
        [TestCase(ComponentType.StandardModule, "Main", "AutoClose", "Word")]
        [TestCase(ComponentType.StandardModule, "Main", "AutoExit", "Word")]
        [Category("Inspections")]
        public void ProcedureNotUsed_NoResultForHostSpecificAutoMacros(ComponentType componentType, string macroName, string moduleName, string hostName)
        {
            var inputCode =
                $@"Private Sub {macroName}()
End Sub";

            var vbe = MockVbeBuilder.BuildFromModules((moduleName, inputCode, componentType));
            vbe.Setup(v => v.HostApplication().AutoMacroIdentifiers).Returns(new []
            {
                new HostAutoMacro(new[] {componentType}, true, moduleName, macroName)
            });

            Assert.AreEqual(0, InspectionResults(vbe.Object).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ProcedureNotUsed
Private Sub Foo()
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureNotUsed_ExcelHotkeyAnnotation_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@ExcelHotkey X
Private Sub Foo()
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureNotUsed_EntryPointAnnotation_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@EntryPoint
Private Sub Foo()
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5490
        [TestCase(@"Name = ""Bizz""", 0)]
        [TestCase(@"mName = ""Bizz""", 1)]
        [Category("Inspections")]
        public void PropertyLet(string assignmentCode, int expectedResults)
        {
            var inputCode =
$@"
Private mName As String

Private Sub Class_Initialize()
    {assignmentCode}
End Sub

Private Property Let Name(ByVal value As String)
    mName = value
End Property
";

            var modules = new(string, string, ComponentType)[]
            {
                (MockVbeBuilder.TestModuleName, inputCode, ComponentType.ClassModule),
            };

            Assert.AreEqual(expectedResults, InspectionResultsForModules(modules).Count(result => result.Target.DeclarationType.HasFlag(DeclarationType.Procedure)));
        }

        [Test]
        [Category("Inspections")]
        public void RecursiveReferenceOnly_ReturnsResult()
        {
            var inputCode =
$@"
Private mName As String

Private Property Let Name(ByVal value As String)
    mName = value
    If Len(mName) > 10 Then
        Name = Left(mName, 8)
    End If
End Property
";

            var modules = new(string, string, ComponentType)[]
            {
                (MockVbeBuilder.TestModuleName, inputCode, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count(result => result.Target.DeclarationType.HasFlag(DeclarationType.Procedure)));
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ProcedureNotUsedInspection(null);

            Assert.AreEqual(nameof(ProcedureNotUsedInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ProcedureNotUsedInspection(state);
        }
    }
}
