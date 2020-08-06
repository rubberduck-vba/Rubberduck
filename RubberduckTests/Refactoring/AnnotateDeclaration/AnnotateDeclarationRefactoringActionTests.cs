using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.Refactorings.Exceptions;

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

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_NoAttributeAnnotation_AsIfNotSet()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@Ignore ProcedureNotUsed
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Single();
                var annotation = new IgnoreAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Inspection, "ProcedureNotUsed")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_AttributeNotThere_AddsAttribute_Module()
        {
            const string code = @"Attribute VB_Exposed = False
Public Sub Foo()
End Sub
";
            const string expectedCode = @"Attribute VB_Exposed = False
Attribute VB_Description = ""MyDesc""
'@ModuleDescription ""MyDesc""
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var annotation = new ModuleDescriptionAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyDesc")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_AttributeNotThere_AddsAttribute_Member()
        {
            const string code = @"
Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@Description ""MyDesc""
Public Sub Foo()
Attribute Foo.VB_Description = ""MyDesc""
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Single();
                var annotation = new DescriptionAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyDesc")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_AttributeNotThere_AddsAttribute_ModuleVariable()
        {
            const string code = @"
Public MyVariable As Variant

Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@VariableDescription ""MyDesc""
Public MyVariable As Variant
Attribute MyVariable.VB_VarDescription = ""MyDesc""

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .Single();
                var annotation = new VariableDescriptionAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyDesc")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_NoAttributeContext_LocalVariable_Throws()
        {
            const string code = @"

Public Sub Foo()
    Dim MyVariable As Variant
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .Single();
                var annotation = new VariableDescriptionAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyDesc")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            Assert.Throws<AttributeRewriteSessionNotSupportedException>(() => RefactoredCode(code, modelBuilder));
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_AttributeAlreadyThere_AdjustsAttribute_Module()
        {
            const string code = @"Attribute VB_Exposed = False
Attribute VB_PredeclaredId = False
Public Sub Foo()
End Sub
";
            const string expectedCode = @"Attribute VB_Exposed = False
Attribute VB_PredeclaredId = True
'@PredeclaredId
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var annotation = new PredeclaredIdAnnotation();
                var arguments = new List<TypedAnnotationArgument>();

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_AttributeAlreadyThere_AdjustsAttribute_Member()
        {
            const string code = @"
Public Sub Foo()
Attribute Foo.VB_Description = ""NotMyDesc""
End Sub
";
            const string expectedCode = @"
'@Description ""MyDesc""
Public Sub Foo()
Attribute Foo.VB_Description = ""MyDesc""
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Single();
                var annotation = new DescriptionAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyDesc")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_AttributeNotThere_AdjustsAttribute_ModuleVariable()
        {
            const string code = @"
Public MyVariable As Variant
Attribute MyVariable.VB_VarDescription = ""NotMyDesc""

Public Sub Foo()
End Sub
";
            const string expectedCode = @"
'@VariableDescription ""MyDesc""
Public MyVariable As Variant
Attribute MyVariable.VB_VarDescription = ""MyDesc""

Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .Single();
                var annotation = new VariableDescriptionAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyDesc")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_DifferentiatesBetweenExtKeys_Add()
        {
            const string code = @"Attribute VB_Ext_Key = ""MyFirstKey"", ""MyFirstValue""
Attribute VB_Ext_Key = ""MySecondKey"", ""MySecondValue""
Attribute VB_Ext_Key = ""MyThirdKey"", ""MyThirdValue""
Public Sub Foo()
End Sub
";
            const string expectedCode = @"Attribute VB_Ext_Key = ""MyFirstKey"", ""MyFirstValue""
Attribute VB_Ext_Key = ""MySecondKey"", ""MySecondValue""
Attribute VB_Ext_Key = ""MyThirdKey"", ""MyThirdValue""
Attribute VB_Ext_Key = ""MyNewKey"", ""MyNewValue""
'@ModuleAttribute VB_Ext_Key, ""MyNewKey"", ""MyNewValue""
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var annotation = new ModuleAttributeAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Attribute, "VB_Ext_Key"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyNewKey"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyNewValue")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_DifferentiatesBetweenExtKeys_Adjust()
        {
            const string code = @"Attribute VB_Ext_Key = ""MyFirstKey"", ""MyFirstValue""
Attribute VB_Ext_Key = ""MySecondKey"", ""MySecondValue""
Attribute VB_Ext_Key = ""MyThirdKey"", ""MyThirdValue""
Public Sub Foo()
End Sub
";
            const string expectedCode = @"Attribute VB_Ext_Key = ""MyFirstKey"", ""MyFirstValue""
Attribute VB_Ext_Key = ""MySecondKey"", ""MyNewValue""
Attribute VB_Ext_Key = ""MyThirdKey"", ""MyThirdValue""
'@ModuleAttribute VB_Ext_Key, ""MySecondKey"", ""MyNewValue""
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var annotation = new ModuleAttributeAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Attribute, "VB_Ext_Key"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MySecondKey"),
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyNewValue")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_AdjustAttributeSet_WorksWithExistingAnnotation_Module()
        {
            const string code = @"Attribute VB_Exposed = False
'@Folder ""MyFolder""
'@DefaultMember
Public Sub Foo()
End Sub
";
            const string expectedCode = @"Attribute VB_Exposed = False
Attribute VB_Description = ""MyDesc""
'@ModuleDescription ""MyDesc""
'@Folder ""MyFolder""
'@DefaultMember
Public Sub Foo()
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var annotation = new ModuleDescriptionAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyDesc")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        [Test]
        [Category("Refactorings")]
        public void AnnotateDeclarationRefactoringAction_WorksWithExistingAnnotation_Member()
        {
            const string code = @"Attribute VB_Exposed = False
'@Folder ""MyFolder""
'@DefaultMember
Public Sub Foo()
End Sub
";
            const string expectedCode = @"Attribute VB_Exposed = False
'@Folder ""MyFolder""
'@DefaultMember
'@Description ""MyDesc""
Public Sub Foo()
Attribute Foo.VB_Description = ""MyDesc""
End Sub
";
            Func<RubberduckParserState, AnnotateDeclarationModel> modelBuilder = (state) =>
            {
                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Single();
                var annotation = new DescriptionAnnotation();
                var arguments = new List<TypedAnnotationArgument>
                {
                    new TypedAnnotationArgument(AnnotationArgumentType.Text, "MyDesc")
                };

                return new AnnotateDeclarationModel(declaration, annotation, arguments, true);
            };

            var refactoredCode = RefactoredCode(code, modelBuilder);

            Assert.AreEqual(expectedCode, refactoredCode);
        }

        protected override IRefactoringAction<AnnotateDeclarationModel> TestBaseRefactoring(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            var annotationUpdater = new AnnotationUpdater(state);
            var attributesUpdater = new AttributesUpdater(state);
            return new AnnotateDeclarationRefactoringAction(rewritingManager, annotationUpdater, attributesUpdater);
        }
    }
}