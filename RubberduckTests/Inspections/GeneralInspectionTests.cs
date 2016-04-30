using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class GeneralInspectionTests
    {
        [TestMethod]
        public void InspectionNameStringsExist()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => type.BaseType == typeof(InspectionBase))
                          .Where(i => string.IsNullOrEmpty(InspectionsUI.ResourceManager.GetString(i.Name + "Name")))
                          .Select(i => i.Name);

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }

        [TestMethod]
        public void InspectionMetaStringsExist()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => type.BaseType == typeof(InspectionBase))
                          .Where(i => string.IsNullOrEmpty(InspectionsUI.ResourceManager.GetString(i.Name + "Meta")))
                          .Select(i => i.Name);

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }

        [TestMethod]
        public void InspectionResultFormatStringsExist()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => type.BaseType == typeof(InspectionBase))
                          .Where(i => string.IsNullOrEmpty(InspectionsUI.ResourceManager.GetString(i.Name + "ResultFormat")))
                          .Select(i => i.Name);

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }

        [TestMethod]
        public void InspectionNameStrings_AreNotFormatted()
        {
            var inspections = typeof(InspectionBase).Assembly.GetTypes()
                          .Where(type => type.BaseType == typeof(InspectionBase))
                          .Where(i =>
                          {
                              var value = InspectionsUI.ResourceManager.GetString(i.Name + "Name");
                              return !string.IsNullOrEmpty(value) && value.Contains("{0}");
                          })
                          .Select(i => i.Name);

            Assert.IsFalse(inspections.Any(), string.Join(Environment.NewLine, inspections));
        }
    }
}