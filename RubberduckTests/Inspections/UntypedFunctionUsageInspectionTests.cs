using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class UntypedFunctionUsageInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void UntypedFunctionUsage_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim str As String
    str = Left(""test"", 1)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 1, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            GetBuiltInDeclarations().ForEach(d => parser.State.AddDeclaration(d));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UntypedFunctionUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UntypedFunctionUsage_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim str As String
    str = Left$(""test"", 1)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            GetBuiltInDeclarations().ForEach(d => parser.State.AddDeclaration(d));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UntypedFunctionUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UntypedFunctionUsage_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim str As String

    '@Ignore UntypedFunctionUsage
    str = Left(""test"", 1)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 1, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            GetBuiltInDeclarations().ForEach(d => parser.State.AddDeclaration(d));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UntypedFunctionUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UntypedFunctionUsage_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
    Dim str As String
    str = Left(""test"", 1)
End Sub";

            const string expectedCode =
@"Sub Foo()
    Dim str As String
    str = Left$(""test"", 1)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 1, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var component = project.Object.VBComponents[0];
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            GetBuiltInDeclarations().ForEach(d => parser.State.AddDeclaration(d));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UntypedFunctionUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            new UntypedFunctionUsageQuickFix(parser.State).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, parser.State.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UntypedFunctionUsage_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
    Dim str As String
    str = Left(""test"", 1)
End Sub";

            const string expectedCode =
@"Sub Foo()
    Dim str As String
'@Ignore UntypedFunctionUsage
    str = Left(""test"", 1)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 1, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var component = project.Object.VBComponents[0];
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            GetBuiltInDeclarations().ForEach(d => parser.State.AddDeclaration(d));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new UntypedFunctionUsageInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(parser.State, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, parser.State.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new UntypedFunctionUsageInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "UntypedFunctionUsageInspection";
            var inspection = new UntypedFunctionUsageInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private List<Declaration> GetBuiltInDeclarations()
        {
            var vbaDeclaration = new ProjectDeclaration(
                new QualifiedMemberName(new QualifiedModuleName("VBA", MockVbeBuilder.LibraryPathVBA, "VBA"), "VBA"),
                "VBA",
                false, null);

            var conversionModule = new ProceduralModuleDeclaration(
                new QualifiedMemberName(new QualifiedModuleName("VBA", MockVbeBuilder.LibraryPathVBA, "Conversion"), "Conversion"),
                vbaDeclaration,
                "Conversion",
                false,
                new List<IAnnotation>(),
                new Attributes());

            var fileSystemModule = new ProceduralModuleDeclaration(
                new QualifiedMemberName(new QualifiedModuleName("VBA", MockVbeBuilder.LibraryPathVBA, "FileSystem"), "FileSystem"),
                vbaDeclaration,
                "FileSystem",
                false,
                new List<IAnnotation>(),
                new Attributes());

            var interactionModule = new ProceduralModuleDeclaration(
                new QualifiedMemberName(new QualifiedModuleName("VBA", MockVbeBuilder.LibraryPathVBA, "Interaction"), "Interaction"),
                vbaDeclaration,
                "Interaction",
                false,
                new List<IAnnotation>(),
                new Attributes());

            var stringsModule = new ProceduralModuleDeclaration(
                new QualifiedMemberName(new QualifiedModuleName("VBA", MockVbeBuilder.LibraryPathVBA, "Strings"), "Strings"),
                vbaDeclaration,
                "Strings",
                false,
                new List<IAnnotation>(),
                new Attributes());

            var dateTimeModule = new ProceduralModuleDeclaration(
                new QualifiedMemberName(new QualifiedModuleName("VBA", MockVbeBuilder.LibraryPathVBA, "DateTime"), "DateTime"),
                vbaDeclaration,
                "Strings",
                false,
                new List<IAnnotation>(),
                new Attributes());

            var hiddenModule = new ProceduralModuleDeclaration(
                new QualifiedMemberName(new QualifiedModuleName("VBA", MockVbeBuilder.LibraryPathVBA, "_HiddenModule"), "_HiddenModule"),
                vbaDeclaration,
                "_HiddenModule",
                false,
                new List<IAnnotation>(),
                new Attributes());


            var commandFunction = new FunctionDeclaration(
                new QualifiedMemberName(interactionModule.QualifiedName.QualifiedModuleName, "_B_var_Command"),
                interactionModule,
                interactionModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var environFunction = new FunctionDeclaration(
                new QualifiedMemberName(interactionModule.QualifiedName.QualifiedModuleName, "_B_var_Environ"),
                interactionModule,
                interactionModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var rtrimFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_RTrim"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var chrFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_Chr"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var formatFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_Format"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var firstFormatParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "Expression"),
                formatFunction,
                "Variant",
                null,
                null,
                false,
                true);

            var secondFormatParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "Format"),
                formatFunction,
                "Variant",
                null,
                null,
                true,
                true);

            var thirdFormatParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "FirstDayOfWeek"),
                formatFunction,
                "VbDayOfWeek",
                null,
                null,
                true,
                false);

            formatFunction.AddParameter(firstFormatParam);
            formatFunction.AddParameter(secondFormatParam);
            formatFunction.AddParameter(thirdFormatParam);

            var rightFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_Right"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var firstRightParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "String"),
                rightFunction,
                "Variant",
                null,
                null,
                false,
                true);

            rightFunction.AddParameter(firstRightParam);

            var lcaseFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_LCase"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var leftbFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_LeftB"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var firstLeftBParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "String"),
                leftbFunction,
                "Variant",
                null,
                null,
                false,
                true);

            leftbFunction.AddParameter(firstLeftBParam);

            var chrwFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_ChrW"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var leftFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_Left"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var firstLeftParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "String"),
                leftFunction,
                "Variant",
                null,
                null,
                false,
                true);

            leftFunction.AddParameter(firstLeftParam);

            var rightbFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_RightB"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var firstRightBParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "String"),
                rightbFunction,
                "Variant",
                null,
                null,
                false,
                true);

            rightbFunction.AddParameter(firstRightBParam);

            var midbFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_MidB"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var firstMidBParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "String"),
                midbFunction,
                "Variant",
                null,
                null,
                false,
                true);

            var secondMidBParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "Start"),
                midbFunction,
                "Long",
                null,
                null,
                false,
                false);

            midbFunction.AddParameter(firstMidBParam);
            midbFunction.AddParameter(secondMidBParam);

            var ucaseFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_UCase"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var trimFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_Trim"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var ltrimFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_LTrim"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var midFunction = new FunctionDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "_B_var_Mid"),
                stringsModule,
                stringsModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var firstMidParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "String"),
                midFunction,
                "Variant",
                null,
                null,
                false,
                true);

            var secondMidParam = new ParameterDeclaration(
                new QualifiedMemberName(stringsModule.QualifiedName.QualifiedModuleName, "Start"),
                midFunction,
                "Long",
                null,
                null,
                false,
                false);

            midFunction.AddParameter(firstMidParam);
            midFunction.AddParameter(secondMidParam);

            var hexFunction = new FunctionDeclaration(
                new QualifiedMemberName(conversionModule.QualifiedName.QualifiedModuleName, "_B_var_Hex"),
                conversionModule,
                conversionModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var octFunction = new FunctionDeclaration(
                new QualifiedMemberName(conversionModule.QualifiedName.QualifiedModuleName, "_B_var_Oct"),
                conversionModule,
                conversionModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var errorFunction = new FunctionDeclaration(
                new QualifiedMemberName(conversionModule.QualifiedName.QualifiedModuleName, "_B_var_Error"),
                conversionModule,
                conversionModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var strFunction = new FunctionDeclaration(
                new QualifiedMemberName(conversionModule.QualifiedName.QualifiedModuleName, "_B_var_Str"),
                conversionModule,
                conversionModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var curDirFunction = new FunctionDeclaration(
                new QualifiedMemberName(fileSystemModule.QualifiedName.QualifiedModuleName, "_B_var_CurDir"),
                fileSystemModule,
                fileSystemModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var datePropertyGet = new PropertyGetDeclaration(
                new QualifiedMemberName(dateTimeModule.QualifiedName.QualifiedModuleName, "Date"),
                dateTimeModule,
                dateTimeModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());


            var timePropertyGet = new PropertyGetDeclaration(
                new QualifiedMemberName(dateTimeModule.QualifiedName.QualifiedModuleName, "Time"),
                dateTimeModule,
                dateTimeModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var inputbFunction = new FunctionDeclaration(
                new QualifiedMemberName(hiddenModule.QualifiedName.QualifiedModuleName, "_B_var_InputB"),
                hiddenModule,
                hiddenModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var firstInputBParam = new ParameterDeclaration(
                new QualifiedMemberName(hiddenModule.QualifiedName.QualifiedModuleName, "Number"),
                inputbFunction,
                "Long",
                null,
                null,
                false,
                true);

            var secondInputBParam = new ParameterDeclaration(
                new QualifiedMemberName(hiddenModule.QualifiedName.QualifiedModuleName, "FileNumber"),
                inputbFunction,
                "Integer",
                null,
                null,
                false,
                true);

            inputbFunction.AddParameter(firstInputBParam);
            inputbFunction.AddParameter(secondInputBParam);

            var inputFunction = new FunctionDeclaration(
                new QualifiedMemberName(hiddenModule.QualifiedName.QualifiedModuleName, "_B_var_Input"),
                hiddenModule,
                hiddenModule,
                "Variant",
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());

            var firstInputParam = new ParameterDeclaration(
                new QualifiedMemberName(hiddenModule.QualifiedName.QualifiedModuleName, "Number"),
                inputFunction,
                "Long",
                null,
                null,
                false,
                true);

            var secondInputParam = new ParameterDeclaration(
                new QualifiedMemberName(hiddenModule.QualifiedName.QualifiedModuleName, "FileNumber"),
                inputFunction,
                "Integer",
                null,
                null,
                false,
                true);

            inputFunction.AddParameter(firstInputParam);
            inputFunction.AddParameter(secondInputParam);

            return new List<Declaration>
            {
                vbaDeclaration,
                conversionModule,
                fileSystemModule,
                interactionModule,
                stringsModule,
                dateTimeModule,
                hiddenModule,
                commandFunction,
                environFunction,
                rtrimFunction,
                chrFunction,
                formatFunction,
                firstFormatParam,
                secondFormatParam,
                thirdFormatParam,
                rightFunction,
                firstRightParam,
                lcaseFunction,
                leftbFunction,
                firstLeftBParam,
                chrwFunction,
                leftFunction,
                firstLeftParam,
                rightbFunction,
                firstRightBParam,
                midbFunction,
                firstMidBParam,
                secondMidBParam,
                ucaseFunction,
                trimFunction,
                ltrimFunction,
                midFunction,
                firstMidParam,
                secondMidParam,
                hexFunction,
                octFunction,
                errorFunction,
                strFunction,
                curDirFunction,
                datePropertyGet,
                timePropertyGet,
                inputbFunction,
                firstInputBParam,
                secondInputBParam,
                inputFunction,
                firstInputParam,
                secondInputParam
            };
        }
    }
}
