using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Resources;
using Rubberduck.Resources.Inspections;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class GeneralInspectionTests
    {
        [SetUp]
        public void InitResources()
        {
            // ensure resources are using an invariant culture.
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            Thread.CurrentThread.CurrentUICulture = Thread.CurrentThread.CurrentCulture;
            InspectionsUI.Culture = Thread.CurrentThread.CurrentUICulture;
            InspectionNames.Culture = Thread.CurrentThread.CurrentUICulture;
            RubberduckUI.Culture = Thread.CurrentThread.CurrentUICulture;
        }

        private static List<Type> GetAllBaseTypes(Type type)
        {
            var baseTypes = new List<Type>();

            var baseType = type.BaseType;
            while (baseType != null)
            {
                baseTypes.Add(baseType);
                baseType = baseType.BaseType;
            }

            return baseTypes;
        }

        [Test]
        [Category("Inspections")]
        public void InspectionNameStringsExist()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => GetAllBaseTypes(type).Contains(typeof(InspectionBase)) && !type.IsAbstract)
                          .Where(i => string.IsNullOrWhiteSpace(InspectionNames.ResourceManager.GetString(i.Name)))
                          .Select(i => i.Name)
                          .ToList();

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }

        [Test]
        [Category("Inspections")]
        public void InspectionNameStringsAreUnique()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                .Where(type => GetAllBaseTypes(type).Contains(typeof(InspectionBase)) && !type.IsAbstract)
                .Select(i => i.Name)
                .ToList();

            Assert.AreEqual(inspections.Count, inspections.Distinct().Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionMetaStringsExist()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => GetAllBaseTypes(type).Contains(typeof(InspectionBase)) && !type.IsAbstract)
                          .Where(i => string.IsNullOrWhiteSpace(InspectionInfo.ResourceManager.GetString(i.Name)))
                          .Select(i => i.Name)
                          .ToList();
            
            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }

        [Test]
        [Category("Inspections")]
        [Category("Inspections")]
        public void InspectionResultFormatStringsExist()
        {
            var inspectionsWithSharedResultFormat = new List<string>
            {
                typeof(UDTMemberNotUsedInspection).Name,
                typeof(ConstantNotUsedInspection).Name,
                typeof(ParameterNotUsedInspection).Name,
                typeof(ProcedureNotUsedInspection).Name,
                typeof(VariableNotUsedInspection).Name,
                typeof(LineLabelNotUsedInspection).Name,
                typeof(UseMeaningfulNameInspection).Name,
                typeof(HungarianNotationInspection).Name,
                typeof(AnnotationInIncompatibleComponentTypeInspection).Name
            };

            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => GetAllBaseTypes(type).Contains(typeof(InspectionBase)) && !type.IsAbstract)
                          .Where(i => !inspectionsWithSharedResultFormat.Contains(i.Name) &&
                                      string.IsNullOrWhiteSpace(InspectionResults.ResourceManager.GetString(i.Name)))
                          .Select(i => i.Name)
                          .ToList();

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }

        [Test]
        [Category("Inspections")]
        public void InspectionNameStrings_AreNotFormatted()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => GetAllBaseTypes(type).Contains(typeof(InspectionBase)))
                          .Where(i =>
                          {
                              var value = InspectionsUI.ResourceManager.GetString(i.Name + "Name");
                              return !string.IsNullOrWhiteSpace(value) && value.Contains("{0}");
                          })
                          .Select(i => i.Name)
                          .ToList();

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }
    }
}
