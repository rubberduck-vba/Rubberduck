using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;

namespace RubberduckTests.Inspections 
{
    class ExcessiveInterfaceMembersInspectionTests : InspectionTestsBase 
    {
        private static List<(string, string, ComponentType)> modules = new List<(string, string, ComponentType)> {
            (null, null, ComponentType.ClassModule),
            ("Class2","Implements Class1", ComponentType.ClassModule) };

        private int AssessCode(string testCode) 
        {
            modules[0] = ("Class1", testCode, ComponentType.ClassModule);
            return InspectionResultsForModules(modules, ReferenceLibrary.Scripting).Count();
        }

        [Test]
        [Category("Inspections")]
        public void ExcessiveInterfaceMembers_Standard() 
        {
            const string testCode =

                @"Public Sub Foo1()
                End Sub

                Public Sub Foo2()
                End Sub

                Public Function Foo3()
                End Function";

            Assert.AreEqual(1, AssessCode(testCode));
        }

        [Test]
        [Category("Inspections")]
        public void ExcessiveInterfaceMembers_IgnoreEvent() 
        {
            const string testCode =

                @"Public Event Event1()

                Public Sub Foo1()
                End Sub

                Public Function Foo2()
                End Function";

            Assert.AreEqual(0, AssessCode(testCode));
        }

        [Test]
        [Category("Inspections")]
        public void ExcessiveInterfaceMembers_OnlySetters() 
        {
            const string testCode =

                @"Public Property Let Foo1(bar1 As Integer)
                End Property

                Public Property Set Foo2(bar2 As Object)
                End Property

                Public Function Foo3()
                End Function";

            Assert.AreEqual(1, AssessCode(testCode));
        }

        [Test]
        [Category("Inspections")]
        public void ExcessiveInterfaceMembers_OnlyGetter()
        {
            const string testCode =

                @"Public Property Get Foo1() As Integer
                End Property

                Public Sub Foo2()
                End Sub

                Public Function Foo3()
                End Function";

            Assert.AreEqual(1, AssessCode(testCode));
        }

        [Test]
        [Category("Inspections")]
        public void ExcessiveInterfaceMembers_ReadWriteProperty()
        {
            const string testCode =

                @"Public Property Let Foo1(bar1 As Variant)
                End Property

                Public Property Get Foo1() As Variant
                End Property

                Public Function Foo2()
                End Function";

            Assert.AreEqual(0, AssessCode(testCode));
        }

        private Mock<IConfigurationService<int>> GetInspectionSettings()
        {
            var settings = new Mock<IConfigurationService<int>>();
            settings.Setup(s => s.Read()).Returns(2);
            return settings;
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state) 
        {
            return new ExcessiveInterfaceMembersInspection(state, GetInspectionSettings().Object);
        }
    }
}
