using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using System.Resources;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ResourceFileInspection
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionNamesFoundInResourceFiles()
        {
            var inspectionAssembly = Assembly.GetAssembly(typeof(Rubberduck.Inspections.Concrete.ApplicationWorksheetFunctionInspection)); //TODO: Elegantly fix
            var iInspectionTypes = inspectionAssembly.GetTypes().Where(type => type.GetInterface("IInspection") != null);

            var missingStrings = (from inspection in iInspectionTypes
                                  let unusedDescription = Rubberduck.Parsing.Inspections.Resources.InspectionsUI.ResourceManager.GetString($"{inspection.Name}Name")
                                  where unusedDescription == null
                                  select inspection.Name)
                            .ToList();

            Assert.IsFalse(missingStrings.Any(), $"Missing values: {string.Join(", ", missingStrings)}");
        }
    }
}
