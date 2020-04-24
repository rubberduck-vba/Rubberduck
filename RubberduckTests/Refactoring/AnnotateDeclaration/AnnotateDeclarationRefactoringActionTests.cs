using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AnnotateDeclaration;

namespace RubberduckTests.Refactoring.AnnotateDeclaration
{
    [TestFixture]
    public class AnnotateDeclarationRefactoringActionTests : RefactoringActionTestBase<AnnotateDeclarationModel>
    {
        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_NoArgument()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@Exposed

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single();
                var annotation = new ExposedModuleAnnotation();
                var arguments = new List<TypedAnnotationArgument>();

                return new AnnotateDeclarationModel(module, annotation, arguments);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_TextArgument()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@Folder ""MyNew""""Folder.MySubFolder""

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single();
                var annotation = new FolderAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyNew\"Folder.MySubFolder")
                };

                return new AnnotateDeclarationModel(module, annotation, arguments);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_InspectionNameArgument()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@IgnoreModule AssignmentNotUsed, FunctionReturnValueNotUsed

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single();
                var annotation = new IgnoreModuleAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Inspection, "AssignmentNotUsed"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Inspection, "FunctionReturnValueNotUsed")
                };

                return new AnnotateDeclarationModel(module, annotation, arguments);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AttributeArgument()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@ModuleAttribute VB_Description, ""MyDescription""

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single();
                var annotation = new ModuleAttributeAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Attribute, "VB_Description"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyDescription")
                };

                return new AnnotateDeclarationModel(module, annotation, arguments);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_BooleanArgument_True()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@ModuleAttribute VB_Exposed, True

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single();
                var annotation = new ModuleAttributeAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Attribute, "VB_Exposed"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Boolean, "true")
                };

                return new AnnotateDeclarationModel(module, annotation, arguments);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_BooleanArgument_False()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@ModuleAttribute VB_Exposed, False

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single();
                var annotation = new ModuleAttributeAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Attribute, "VB_Exposed"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Boolean, "faLse")
                };

                return new AnnotateDeclarationModel(module, annotation, arguments);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_BooleanArgument_Undefined()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"'@ModuleAttribute VB_Exposed, NOT_A_BOOLEAN

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .Single();
                var annotation = new ModuleAttributeAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Attribute, "VB_Exposed"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Boolean, "aefefef")
                };

                return new AnnotateDeclarationModel(module, annotation, arguments);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_NumberArgument()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@MemberAttribute VB_UserMemId, 0
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var module = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Single();
                var annotation = new MemberAttributeAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Attribute, "VB_UserMemId"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Number, "0")
                };

                return new AnnotateDeclarationModel(module, annotation, arguments);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }


        protected override IRefactoringAction<AnnotateDeclarationModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var annotationUpdater = new AnnotationUpdater();
            return new AnnotateDeclarationRefactoringAction(rewritingManager, annotationUpdater);
        }
    }
}