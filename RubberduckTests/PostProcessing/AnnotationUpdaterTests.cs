using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.PostProcessing
{
    [TestFixture]
    public class AnnotationUpdaterTests
    {
        [Test]
        [Category("AnnotationUpdater")]
        public void AddAnnotationAddsMemberAnnotationRightAboveMember()
        {
            const string inputCode =
                @"
Private Sub FooBar() 
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"
Private Sub FooBar() 
End Sub


'@Obsolete
    '@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            var annotationToAdd = new MemberAttributeAnnotation();
            var annotationValues = new List<string> { "VB_Ext_Key", "\"Key\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.AddAnnotation(rewriteSession, fooDeclaration, annotationToAdd, annotationValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void AddAnnotationAddsModuleAnnotationAboveTheFirstLineForCodePaneRewriteSession()
        {
            const string inputCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            var annotationToAdd = new ModuleAttributeAnnotation();
            var annotationValues = new List<string> { "VB_Ext_Key", "\"Key\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.AddAnnotation(rewriteSession, moduleDeclaration, annotationToAdd, annotationValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void AddAnnotationAddsModuleAnnotationBelowTheLastAttributeForAttributeRewriteSession()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            var annotationToAdd = new ModuleAttributeAnnotation();
            var annotationValues = new List<string> { "VB_Ext_Key", "\"Key\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode, ComponentType.ClassModule, CodeKind.AttributesCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ClassModule)
                    .First();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.AddAnnotation(rewriteSession, moduleDeclaration, annotationToAdd, annotationValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void AddAnnotationDoesNotAddModuleAnnotationsToMembers()
        {
            const string inputCode =
                @"
Private Sub FooBar() 
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"
Private Sub FooBar() 
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            var annotationToAdd = new ModuleAttributeAnnotation();
            var annotationValues = new List<string> { "VB_Ext_Key", "\"Key\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.AddAnnotation(rewriteSession, fooDeclaration, annotationToAdd, annotationValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void AddAnnotationDoesNotAddMemberAnnotationsToModules()
        {
            const string inputCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            var annotationToAdd = new MemberAttributeAnnotation();
            var annotationValues = new List<string> { "VB_Ext_Key", "\"Key\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.AddAnnotation(rewriteSession, moduleDeclaration, annotationToAdd, annotationValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void AddAnnotationAddsIdentifierReferenceAnnotationsRightAboveTheReference()
        {
            const string inputCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar()
    Dim bar As Variant
    bar = Foo(""x"")
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar()
    Dim bar As Variant
    '@Ignore ObsoleteMemberUsage
    bar = Foo(""x"")
End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            var annotationToAdd = new IgnoreAnnotation();
            var annotationValues = new List<string> { "ObsoleteMemberUsage" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooReference = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo")
                    .References.First();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.AddAnnotation(rewriteSession, fooReference, annotationToAdd, annotationValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void AddAnnotationOnTheFirstLineIgnoresIndentation()
        {
            const string inputCode =
                @"  Private Sub FooBar() 
  End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"'@Obsolete
  Private Sub FooBar() 
  End Sub


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            var annotationToAdd = new ObsoleteAnnotation();

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooBarDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "FooBar");
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.AddAnnotation(rewriteSession, fooBarDeclaration, annotationToAdd);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void AddAnnotationNotOnFirstPhysicalLineOfALogicalLineDoesNothing()
        {
            const string inputCode =
                @"
Private fooBar As Long, _
baz As Variant


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"
Private fooBar As Long, _
baz As Variant


'@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            var annotationToAdd = new ObsoleteAnnotation();

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var bazDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .First(decl => decl.IdentifierName == "baz");
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.AddAnnotation(rewriteSession, bazDeclaration, annotationToAdd);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void RemoveAnnotationLeavesCommentOnSameLine()
        {
            const string inputCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


 '@Obsolete :It is obsolete.
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


 'It is obsolete.
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var annotationToRemove = fooDeclaration.Annotations.First();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.RemoveAnnotation(rewriteSession, annotationToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void RemoveAnnotationWorksForFirstAnnotation()
        {
            const string inputCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


 '@Obsolete @Description ""Desc"" @DefaultMember
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


 '@Description ""Desc"" @DefaultMember
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var annotationToRemove = fooDeclaration.Annotations.First();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.RemoveAnnotation(rewriteSession, annotationToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void RemoveAnnotationWorksForLastAnnotation()
        {
            const string inputCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


 '@Obsolete @Description ""Desc"" @DefaultMember
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


 '@Obsolete @Description ""Desc"" 
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var annotationToRemove = fooDeclaration.Annotations.Last();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.RemoveAnnotation(rewriteSession, annotationToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void RemoveAnnotationForStandaloneAnnotationRemovesEntireLine()
        {
            const string inputCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub

     'Strange indent comment
 '@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub

     'Strange indent comment
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var annotationToRemove = fooDeclaration.Annotations.First();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.RemoveAnnotation(rewriteSession, annotationToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void RemoveAnnotationForStandaloneAnnotationRemovesEntireLineOnFirstLine()
        {
            const string inputCode =
                @" '@PredeclaredId
Option Explicit

Private Sub FooBar() 
End Sub

     'Strange indent comment
 '@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"Option Explicit

Private Sub FooBar() 
End Sub

     'Strange indent comment
 '@Obsolete
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var annotationToRemove = moduleDeclaration.Annotations.First();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.RemoveAnnotation(rewriteSession, annotationToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void RemoveAnnotationForStandaloneAnnotationRemovesEntireLineOnLastLine()
        {
            const string inputCode =
                @"
Option Explicit
   'Strange indent comment
 '@PredeclaredId";

            const string expectedCode =
                @"
Option Explicit
   'Strange indent comment";
            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var annotationToRemove = moduleDeclaration.Annotations.First();
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.RemoveAnnotation(rewriteSession, annotationToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void RemoveAnnotationsWorks()
        {
            const string inputCode =
                @"'@ModuleAttribute SomeAttribute
Option Explicit
'@ModuleAttribute SomeOtherAttribute @ModuleAttribute YetAnotherAttribute
   'Strange indent comment
    '@ModuleAttribute YetYetAnotherAttribute @Exposed
  '@Folder ""Folder""
 '@PredeclaredId @ModuleDescription ""Desc""";

            const string expectedCode =
                @"Option Explicit
   'Strange indent comment
    '@Exposed
";
            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var annotationsToRemove = moduleDeclaration.Annotations.Where(pta => !(pta.Annotation is ExposedModuleAnnotation));
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.RemoveAnnotations(rewriteSession, annotationsToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AnnotationUpdater")]
        public void UpdateAnnotationReplacesEntireAnnotation()
        {
            const string inputCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


 '@Obsolete @Description ""Desc"" @DefaultMember
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";

            const string expectedCode =
                @"'@PredeclaredId
Option Explicit
'@Folder ""folder""

Private Sub FooBar() 
End Sub


 '@Obsolete @MemberAttribute VB_ExtKey, ""Key"", ""Value"" @DefaultMember
    Public Sub Foo(bar As String)
        bar = vbNullString
    End Sub
";
            var newAnnotation = new MemberAttributeAnnotation();
            var newAnnotationValues = new List<string> { "VB_ExtKey", "\"Key\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var annotationToUpdate = fooDeclaration.Annotations.First(pta => pta.Annotation is DescriptionAnnotation);
                var annotationUpdater = new AnnotationUpdater(state);

                annotationUpdater.UpdateAnnotation(rewriteSession, annotationToUpdate, newAnnotation, newAnnotationValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        private (IVBComponent component, IExecutableRewriteSession rewriteSession, RubberduckParserState state) TestSetup(string inputCode, ComponentType componentType = ComponentType.StandardModule, CodeKind codeKind = CodeKind.CodePaneCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, componentType, out var component).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            var rewriteSession = codeKind == CodeKind.AttributesCode
                ? rewritingManager.CheckOutAttributesSession()
                : rewritingManager.CheckOutCodePaneSession();
            return (component, rewriteSession, state);
        }
    }
}