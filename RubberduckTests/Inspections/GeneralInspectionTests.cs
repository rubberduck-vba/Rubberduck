using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Castle.Core.Internal;
using NUnit.Framework;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
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
            Rubberduck.Resources.Inspections.InspectionsUI.Culture = Thread.CurrentThread.CurrentUICulture;
            Rubberduck.Resources.Inspections.InspectionNames.Culture = Thread.CurrentThread.CurrentUICulture;
            Rubberduck.Resources.Inspections.InspectionInfo.Culture = Thread.CurrentThread.CurrentUICulture;
            Rubberduck.Resources.Inspections.InspectionResults.Culture = Thread.CurrentThread.CurrentUICulture;
            Rubberduck.Resources.Inspections.QuickFixes.Culture = Thread.CurrentThread.CurrentUICulture;
            RubberduckUI.Culture = Thread.CurrentThread.CurrentUICulture;
            Rubberduck.Resources.CodeExplorer.Culture = Thread.CurrentThread.CurrentUICulture;
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
                .ToList();
            var resources = inspections
                .ToDictionary(i => i.Name, i => InspectionNames.ResourceManager.GetString(i.Name, CultureInfo.InvariantCulture))
                .ToList();
            var missingKeys = resources
                .Where(r => r.Value.IsNullOrEmpty())
                .Select(i => i.Key)
                .ToList();

            Assert.IsFalse(missingKeys.Any(), string.Join(Environment.NewLine, inspections));
        }

        [Test]
        [Category("Inspections")]
        public void InspectionMetaStringsExist()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => GetAllBaseTypes(type).Contains(typeof(InspectionBase)) && !type.IsAbstract)
                          .Where(i => string.IsNullOrWhiteSpace(InspectionInfo.ResourceManager.GetString(i.Name, CultureInfo.InvariantCulture)))
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
                typeof(ConstantNotUsedInspection).Name,
                typeof(ParameterNotUsedInspection).Name,
                typeof(ProcedureNotUsedInspection).Name,
                typeof(VariableNotUsedInspection).Name,
                typeof(LineLabelNotUsedInspection).Name,
                typeof(UseMeaningfulNameInspection).Name,
                typeof(HungarianNotationInspection).Name
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
                              var value = InspectionNames.ResourceManager.GetString(i.Name);
                              return !string.IsNullOrWhiteSpace(value) && value.Contains("{0}");
                          })
                          .Select(i => i.Name)
                          .ToList();

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }
    }
}
