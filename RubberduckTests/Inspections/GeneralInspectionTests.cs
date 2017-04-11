using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.UI;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class GeneralInspectionTests
    {
        [TestInitialize]
        public void InitResources()
        {
            // ensure resources are using an invariant culture.
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            Thread.CurrentThread.CurrentUICulture = Thread.CurrentThread.CurrentCulture;
            InspectionsUI.Culture = Thread.CurrentThread.CurrentUICulture;
            RubberduckUI.Culture = Thread.CurrentThread.CurrentUICulture;
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionNameStringsExist()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => type.BaseType == typeof(InspectionBase))
                          .Where(i => string.IsNullOrEmpty(InspectionsUI.ResourceManager.GetString(i.Name + "Name")))
                          .Select(i => i.Name)
                          .ToList();

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionMetaStringsExist()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => type.BaseType == typeof(InspectionBase))
                          .Where(i => string.IsNullOrEmpty(InspectionsUI.ResourceManager.GetString(i.Name + "Meta")))
                          .Select(i => i.Name)
                          .ToList();

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }

        [TestMethod]
        [TestCategory("Inspections")]
        [TestCategory("Inspections")]
        public void InspectionResultFormatStringsExist()
        {
            var inspectionsWithSharedResultFormat = new List<string>
            {
                typeof(ConstantNotUsedInspection).Name,
                typeof(ParameterNotUsedInspection).Name,
                typeof(ProcedureNotUsedInspection).Name,
                typeof(VariableNotUsedInspection).Name,
                typeof(UseMeaningfulNameInspection).Name,
                typeof(HungarianNotationInspection).Name
            };

            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => type.BaseType == typeof(InspectionBase))
                          .Where(i => !inspectionsWithSharedResultFormat.Contains(i.Name) &&
                                      string.IsNullOrEmpty(InspectionsUI.ResourceManager.GetString(i.Name + "ResultFormat")))
                          .Select(i => i.Name)
                          .ToList();

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionNameStrings_AreNotFormatted()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => type.BaseType == typeof(InspectionBase))
                          .Where(i =>
                          {
                              var value = InspectionsUI.ResourceManager.GetString(i.Name + "Name");
                              return !string.IsNullOrEmpty(value) && value.Contains("{0}");
                          })
                          .Select(i => i.Name)
                          .ToList();

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }
    }
}
