using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;

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

            var resourceManager = new System.Resources.ResourceManager("InspectionsUI", Assembly.GetAssembly(typeof(Rubberduck.Parsing.Inspections.Resources.InspectionsUI)));

            foreach (var inspection in iInspectionTypes)
            {
                Assert.AreEqual(resourceManager.GetString(inspection.Name + "Name"), inspection.Name);
            }
        }
    }
}
