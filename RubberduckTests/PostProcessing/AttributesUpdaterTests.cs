using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.PostProcessing
{
    [TestFixture]
    public class AttributesUpdaterTests
    {
        [Test]
        [Category("AttributesUpdater")]
        public void AddAttributeAddsMemberAttributeBelowFirstLineOfMember()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
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
Public Sub Foo(bar As String)
Attribute Foo.VB_Description = ""The MyFunc Description""
    bar = vbNullString
End Sub
";
            var attributeToAdd = "Foo.VB_Description";
            var attributeValues = new List<string> {"\"The MyFunc Description\""};

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddAttribute(rewriteSession, fooDeclaration, attributeToAdd, attributeValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddAttributeAddsMemberAttributeBelowOneLineMember()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Public Sub Foo(bar As String) : bar = vbNullString : End Sub
";

            const string expectedCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Public Sub Foo(bar As String) : bar = vbNullString : End Sub
Attribute Foo.VB_Description = ""The MyFunc Description""
";
            var attributeToAdd = "Foo.VB_Description";
            var attributeValues = new List<string> { "\"The MyFunc Description\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddAttribute(rewriteSession, fooDeclaration, attributeToAdd, attributeValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void MultipleAddAttributeWorkForMembers()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
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
Public Sub Foo(bar As String)
Attribute Foo.VB_Description = ""The MyFunc Description""
Attribute Foo.VB_HelpID = 2
    bar = vbNullString
End Sub
";
            var firstAttributeToAdd = "Foo.VB_Description";
            var firstAttributeValues = new List<string> { "\"The MyFunc Description\"" };
            var secondAttributeToAdd = "Foo.VB_HelpID";
            var secondAttributeValues = new List<string> { "2" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddAttribute(rewriteSession, fooDeclaration, firstAttributeToAdd, firstAttributeValues);
                attributesUpdater.AddAttribute(rewriteSession, fooDeclaration, secondAttributeToAdd, secondAttributeValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddAttributeAddsModuleAttributeBelowLastModuleAttribute()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
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
Attribute VB_Exposed = False
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attributeToAdd = "VB_Exposed";
            var attributeValues = new List<string> { "False" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddAttribute(rewriteSession, moduleDeclaration, attributeToAdd, attributeValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void MultipleAddAttributeWorkForModules()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
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
Attribute VB_Exposed = False
Attribute VB_Description = ""Module Description""
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var firstAttributeToAdd = "VB_Exposed";
            var firstAttributeValues = new List<string> { "False" };
            var secondAttributeToAdd = "VB_Description";
            var secondAttributeValues = new List<string> { "\"Module Description\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddAttribute(rewriteSession, moduleDeclaration, firstAttributeToAdd, firstAttributeValues);
                attributesUpdater.AddAttribute(rewriteSession, moduleDeclaration, secondAttributeToAdd, secondAttributeValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        //Should never happen in a real module.
        public void AddAttributeAddsModuleAttributeAtTopOfModuleIfThereAreNoModuleAttributesYet()
        {
            const string inputCode =
                @"Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";

            const string expectedCode =
                @"Attribute VB_Exposed = False
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attributeToAdd = "VB_Exposed";
            var attributeValues = new List<string> { "False" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddAttribute(rewriteSession, moduleDeclaration, attributeToAdd, attributeValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddAttributeDoesNotAddAttributeAlreadyThere()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Exposed = False
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
Attribute VB_Exposed = False
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attributeToAdd = "VB_Exposed";
            var attributeValues = new List<string> { "True" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddAttribute(rewriteSession, moduleDeclaration, attributeToAdd, attributeValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddAttributeAddsVbExtKeyAttributeForDifferentKey()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Ext_Key = ""Key"", ""Value""
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
Attribute VB_Ext_Key = ""Key"", ""Value""
Attribute VB_Ext_Key = ""OtherKey"", ""Value""
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attributeToAdd = "VB_Ext_Key";
            var attributeValues = new List<string> { "\"OtherKey\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddAttribute(rewriteSession, moduleDeclaration, attributeToAdd, attributeValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddAttributeDoesNotAddVbExtKeyAttributeForExistingKey()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Ext_Key = ""Key"", ""Value""
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
Attribute VB_Ext_Key = ""Key"", ""Value""
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attributeToAdd = "VB_Ext_Key";
            var attributeValues = new List<string> { "\"Key\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddAttribute(rewriteSession, moduleDeclaration, attributeToAdd, attributeValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void RemoveAttributeWithoutValuesSpecifiedRemovesAllEntriesForTheAttributeForModules()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Ext_Key = ""Key"", ""Value""
Attribute VB_Ext_Key = ""OtherKey"", ""Value""
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
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attributesToRemove = "VB_Ext_Key";

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.RemoveAttribute(rewriteSession, moduleDeclaration, attributesToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void RemoveAttributeWithoutValuesSpecifiedRemovesAllEntriesForTheAttributeForMembers()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Public Sub Foo(bar As String)
Attribute Foo.VB_Ext_Key = ""Key"", ""Value""
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""Value""
Attribute Foo.VB_Description = ""Desc""
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
Public Sub Foo(bar As String)
Attribute Foo.VB_Description = ""Desc""
    bar = vbNullString
End Sub
";
            var attributesToRemove = "Foo.VB_Ext_Key";

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.RemoveAttribute(rewriteSession, fooDeclaration, attributesToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void RemoveAttributeWithValuesSpecifiedRemovesOnlyEntriesForTheAttributeWithSpecifiedValuesForModules()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Ext_Key = ""Key"", ""Value""
Attribute VB_Ext_Key = ""OtherKey"", ""Value""
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
Attribute VB_Ext_Key = ""OtherKey"", ""Value""
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attributesToRemove = "VB_Ext_Key";
            var valuesToRemove = new List<string> { "\"Key\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.RemoveAttribute(rewriteSession, moduleDeclaration, attributesToRemove, valuesToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void RemoveAttributeWithValuesSpecifiedRemovesOnlyEntriesForTheAttributeWithSpecifiedValuesForMembers()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Public Sub Foo(bar As String)
Attribute Foo.VB_Ext_Key = ""Key"", ""Value""
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""Value""
Attribute Foo.VB_Description = ""Desc""
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
Public Sub Foo(bar As String)
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""Value""
Attribute Foo.VB_Description = ""Desc""
    bar = vbNullString
End Sub
";
            var attributesToRemove = "Foo.VB_Ext_Key";
            var valuesToRemove = new List<string> { "\"Key\"", "\"Value\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.RemoveAttribute(rewriteSession, fooDeclaration, attributesToRemove, valuesToRemove);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void UpdateAttributeWithoutValuesSpecifiedUpdatesFirstEntryForTheAttributeAndRemovesTheRestForModules()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Ext_Key = ""Key"", ""Value""
Attribute VB_Ext_Key = ""OtherKey"", ""Value""
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
Attribute VB_Ext_Key = ""AnotherKey"", ""AnotherValue""
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attributesToRemove = "VB_Ext_Key";
            var newValues = new List<string> { "\"AnotherKey\"", "\"AnotherValue\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.UpdateAttribute(rewriteSession, moduleDeclaration, attributesToRemove, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void UpdateAttributeWithoutValuesSpecifiedUpdatesFirstEntryForTheAttributeAndRemovesTheRestForMembers()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Public Sub Foo(bar As String)
Attribute Foo.VB_Ext_Key = ""Key"", ""Value""
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""Value""
Attribute Foo.VB_Description = ""Desc""
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
Public Sub Foo(bar As String)
Attribute Foo.VB_Ext_Key = ""AnotherKey"", ""AnotherValue""
Attribute Foo.VB_Description = ""Desc""
    bar = vbNullString
End Sub
";
            var attributesToRemove = "Foo.VB_Ext_Key";
            var newValues = new List<string> { "\"AnotherKey\"", "\"AnotherValue\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.UpdateAttribute(rewriteSession, fooDeclaration, attributesToRemove, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void UpdateAttributeWithValuesSpecifiedUpdatesOnlyEntriesForTheAttributeWithSpecifiedValuesForModules()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Ext_Key = ""Key"", ""Value""
Attribute VB_Ext_Key = ""OtherKey"", ""Value""
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
Attribute VB_Ext_Key = ""AnotherKey"", ""AnotherValue""
Attribute VB_Ext_Key = ""OtherKey"", ""Value""
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attributesToRemove = "VB_Ext_Key";
            var oldValues = new List<string> { "\"Key\"", "\"Value\"" };
            var newValues = new List<string> { "\"AnotherKey\"", "\"AnotherValue\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.ProceduralModule)
                    .First();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.UpdateAttribute(rewriteSession, moduleDeclaration, attributesToRemove, newValues, oldValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void UpdateAttributeWithValuesSpecifiedUpdatesOnlyEntriesForTheAttributeWithSpecifiedValuesForMembers()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Public Sub Foo(bar As String)
Attribute Foo.VB_Ext_Key = ""Key"", ""Value""
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""Value""
Attribute Foo.VB_Description = ""Desc""
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
Public Sub Foo(bar As String)
Attribute Foo.VB_Ext_Key = ""AnotherKey"", ""AnotherValue""
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""Value""
Attribute Foo.VB_Description = ""Desc""
    bar = vbNullString
End Sub
";
            var attributesToRemove = "Foo.VB_Ext_Key";
            var oldValues = new List<string> { "\"Key\"", "\"Value\"" };
            var newValues = new List<string> { "\"AnotherKey\"", "\"AnotherValue\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var fooDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .First(decl => decl.IdentifierName == "Foo");
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.UpdateAttribute(rewriteSession, fooDeclaration, attributesToRemove, newValues, oldValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddOrUpdateAttribute_UsualAttribute_NotThere_AddsAttribute_Module()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
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
Attribute VB_Exposed = True
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attribute = "VB_Exposed";
            var newValues = new List<string> { "True" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddOrUpdateAttribute(rewriteSession, moduleDeclaration, attribute, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddOrUpdateAttribute_UsualAttribute_AlreadyThere_UpdatesAttribute_Module()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Exposed = False
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
Attribute VB_Exposed = True
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attribute = "VB_Exposed";
            var newValues = new List<string> { "True" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddOrUpdateAttribute(rewriteSession, moduleDeclaration, attribute, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddOrUpdateAttribute_UsualAttribute_NotThere_AddsAttribute_Member()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
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
Public Sub Foo(bar As String)
Attribute Foo.VB_UserMemId = 0
    bar = vbNullString
End Sub
";
            var attribute = "Foo.VB_UserMemId";
            var newValues = new List<string> { "0" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Single();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddOrUpdateAttribute(rewriteSession, moduleDeclaration, attribute, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddOrUpdateAttribute_UsualAttribute_AlreadyThere_UpdatesAttribute_Member()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Public Sub Foo(bar As String)
Attribute Foo.VB_UserMemId = -4
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
Public Sub Foo(bar As String)
Attribute Foo.VB_UserMemId = 0
    bar = vbNullString
End Sub
";
            var attribute = "Foo.VB_UserMemId";
            var newValues = new List<string> { "0" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Procedure)
                    .Single();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddOrUpdateAttribute(rewriteSession, moduleDeclaration, attribute, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddOrUpdateAttribute_UsualAttribute_NotThere_AddsAttribute_Variable()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False

Public MyVariable As Variant

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

Public MyVariable As Variant
Attribute MyVariable.VB_VarDescription = ""MyDesc""

Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attribute = "MyVariable.VB_VarDescription";
            var newValues = new List<string> { "\"MyDesc\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .Single(declaration => declaration.IdentifierName.Equals("MyVariable"));
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddOrUpdateAttribute(rewriteSession, moduleDeclaration, attribute, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddOrUpdateAttribute_UsualAttribute_AlreadyThere_UpdatesAttribute_Variable()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False

Public MyVariable As Variant
Attribute MyVariable.VB_VarDescription = ""NotMyDesc""

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

Public MyVariable As Variant
Attribute MyVariable.VB_VarDescription = ""MyDesc""

Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attribute = "MyVariable.VB_VarDescription";
            var newValues = new List<string> { "\"MyDesc\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Variable)
                    .Single(declaration => declaration.IdentifierName.Equals("MyVariable"));
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddOrUpdateAttribute(rewriteSession, moduleDeclaration, attribute, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddOrUpdateAttribute_ExtKey_NotThere_AddsAttribute()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
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
Attribute VB_Ext_Key = ""MyKey"", ""MyValue""
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attribute = "VB_Ext_Key";
            var newValues = new List<string> { "\"MyKey\"", "\"MyValue\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddOrUpdateAttribute(rewriteSession, moduleDeclaration, attribute, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddOrUpdateAttribute_ExtKey_KeyNotThere_AddsAttribute()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Ext_Key = ""AnotherKey"", ""MyValuse""
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
Attribute VB_Ext_Key = ""AnotherKey"", ""MyValuse""
Attribute VB_Ext_Key = ""MyKey"", ""MyValue""
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attribute = "VB_Ext_Key";
            var newValues = new List<string> { "\"MyKey\"", "\"MyValue\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddOrUpdateAttribute(rewriteSession, moduleDeclaration, attribute, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("AttributesUpdater")]
        public void AddOrUpdateAttribute_ExtKey_KeyAlreadyThere_UpdatesAttribute()
        {
            const string inputCode =
                @"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""ClassKeys""
Attribute VB_GlobalNameSpace = False
Attribute VB_Ext_Key = ""AnotherKey"", ""MyValue""
Attribute VB_Ext_Key = ""MyKey"", ""AnotherValue""
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
Attribute VB_Ext_Key = ""AnotherKey"", ""MyValue""
Attribute VB_Ext_Key = ""MyKey"", ""MyValue""
Public Sub Foo(bar As String)
    bar = vbNullString
End Sub
";
            var attribute = "VB_Ext_Key";
            var newValues = new List<string> { "\"MyKey\"", "\"MyValue\"" };

            string actualCode;
            var (component, rewriteSession, state) = TestSetup(inputCode);
            using (state)
            {
                var moduleDeclaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Single();
                var attributesUpdater = new AttributesUpdater(state);

                attributesUpdater.AddOrUpdateAttribute(rewriteSession, moduleDeclaration, attribute, newValues);
                rewriteSession.TryRewrite();

                actualCode = component.CodeModule.Content();
            }
            Assert.AreEqual(expectedCode, actualCode);
        }

        private (IVBComponent component, IExecutableRewriteSession rewriteSession, RubberduckParserState state) TestSetup(string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component).Object;
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            return (component, rewritingManager.CheckOutAttributesSession(), state);
        }
    }
}