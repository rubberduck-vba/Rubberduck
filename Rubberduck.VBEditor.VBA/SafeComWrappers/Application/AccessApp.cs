using System.Collections.Generic;
using Microsoft.Office.Interop.Access;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class AccessApp : HostApplicationBase<Microsoft.Office.Interop.Access.Application>
    {
        public AccessApp() : base("Access") { }

        public override HostDocument GetDocument(QualifiedModuleName moduleName)
        {
            if (moduleName.ComponentName.StartsWith("Form_"))
            {
                var name = moduleName.ComponentName.Substring("Form_".Length);
                using (var currentProject = new SafeIDispatchWrapper<_CurrentProject>(Application.CurrentProject))
                using (var allForms = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllForms))
                using (var forms = new SafeIDispatchWrapper<Forms>(Application.Forms))
                using (var accessObject = new SafeIDispatchWrapper<AccessObject>(allForms.Target[name]))
                { 
                    return LoadHostDocument("Access.Form", accessObject, forms);
                }
            }

            if (moduleName.ComponentName.StartsWith("Report_"))
            {
                var name = moduleName.ComponentName.Substring("Report_".Length);
                using (var currentProject = new SafeIDispatchWrapper<_CurrentProject>(Application.CurrentProject))
                using (var allReports = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllReports))
                using (var reports = new SafeIDispatchWrapper<Reports>(Application.Reports))
                using (var accessObject = new SafeIDispatchWrapper<AccessObject>(allReports.Target[name]))
                {
                    return LoadHostDocument("Access.Report", accessObject, reports);
                }
            }

            return null;
        }

        public override IEnumerable<HostDocument> GetDocuments()
        {
            var result = new List<HostDocument>();
            using (var currentProject = new SafeIDispatchWrapper<_CurrentProject>(Application.CurrentProject))
            using (var allForms = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllForms))
            using (var allReports = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllReports))
            using (var forms = new SafeIDispatchWrapper<Forms>(Application.Forms))
            using (var reports = new SafeIDispatchWrapper<Reports>(Application.Reports))
            {
                PopulateList(ref result, "Access.Form", allForms, forms);
                PopulateList(ref result, "Access.Report", allReports, reports);
            }

            return result;
        }

        private void PopulateList(ref List<HostDocument> result, string className, SafeIDispatchWrapper<AllObjects> allObjects, dynamic objects)
        {
            foreach (AccessObject rawAccessObject in allObjects.Target)
            using (var accessObject = new SafeIDispatchWrapper<AccessObject>(rawAccessObject))
            {
                var item = LoadHostDocument(className, accessObject, objects);
                result.Add(item);
            }
        }

        private HostDocument LoadHostDocument(string className, SafeIDispatchWrapper<AccessObject> accessObject, dynamic objects)
        {
            var state = DocumentState.Closed;
            if (!accessObject.Target.IsLoaded)
            {
                return new HostDocument(accessObject.Target.Name, className, state, null);
            }

            state = accessObject.Target.CurrentView == AcCurrentView.acCurViewDesign
                ? DocumentState.DesignView
                : DocumentState.ActiveView;
            return new HostDocument(accessObject.Target.Name, className, state, null);
        }
    }
}
