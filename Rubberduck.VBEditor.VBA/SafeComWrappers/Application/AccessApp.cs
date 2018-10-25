using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Access;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class AccessApp : HostApplicationBase<Microsoft.Office.Interop.Access.Application>
    {
        public AccessApp() : base("Access") { }

        public override IHostDocument GetDocument(QualifiedModuleName moduleName)
        {
            if (moduleName.ComponentName.StartsWith("Form_"))
            {
                var name = moduleName.ComponentName.Substring("Form_".Length);
                _CurrentProject currentProject = null;
                AllObjects allForms = null;
                AccessObject accessObject = null;
                Forms forms = null;
                try
                {
                    currentProject = Application.CurrentProject;
                    forms = Application.Forms;
                    allForms = currentProject.AllForms;
                    accessObject = allForms[name];

                    return LoadHostDocument("Access.Form", accessObject, forms);
                }
                finally
                {
                    if (forms != null) Marshal.ReleaseComObject(forms);
                    if (accessObject != null) Marshal.ReleaseComObject(accessObject);
                    if (allForms != null) Marshal.ReleaseComObject(allForms);
                    if (currentProject != null) Marshal.ReleaseComObject(currentProject);
                }
            }

            if (moduleName.ComponentName.StartsWith("Report_"))
            {
                var name = moduleName.ComponentName.Substring("Report_".Length);
                _CurrentProject currentProject = null;
                AllObjects allForms = null;
                AccessObject accessObject = null;
                Reports reports = null;
                try
                {
                    currentProject = Application.CurrentProject;
                    reports = Application.Reports;
                    allForms = currentProject.AllForms;
                    accessObject = allForms[name];

                    return LoadHostDocument("Access.Report", accessObject, reports);
                }
                finally
                {
                    if (reports != null) Marshal.ReleaseComObject(reports);
                    if (accessObject != null) Marshal.ReleaseComObject(accessObject);
                    if (allForms != null) Marshal.ReleaseComObject(allForms);
                    if (currentProject != null) Marshal.ReleaseComObject(currentProject);
                }
            }

            return null;
        }

        public override IEnumerable<IHostDocument> GetDocuments()
        {
            var result = new List<HostDocument>();
            _CurrentProject currentProject = null;
            AllObjects allObjects = null;
            Forms forms = null;
            Reports reports = null;

            try
            {
                currentProject = Application.CurrentProject;
                allObjects = currentProject.AllForms;
                forms = Application.Forms;

                PopulateList(ref result, "Access.Form", allObjects, forms);

                Marshal.ReleaseComObject(allObjects);

                allObjects = currentProject.AllReports;
                reports = Application.Reports;

                PopulateList(ref result, "Access.Report", allObjects, reports);
            }
            finally
            {
                if (allObjects != null) Marshal.ReleaseComObject(allObjects);
                if (forms != null) Marshal.ReleaseComObject(forms);
                if (reports != null) Marshal.ReleaseComObject(reports);
                if (currentProject != null) Marshal.ReleaseComObject(currentProject);
            }
            
            return result;
        }

        private void PopulateList(ref List<HostDocument> result, string className, AllObjects allObjects, dynamic objects)
        {
            foreach (AccessObject accessObject in allObjects)
            {
                var item = LoadHostDocument(className, accessObject, objects);
                result.Add(item);
            }
        }

        private HostDocument LoadHostDocument(string className, AccessObject accessObject, dynamic objects)
        {
            var state = DocumentState.Closed;
            if (!accessObject.IsLoaded)
            {
                return new HostDocument(accessObject.Name, className, null, state);
            }

            object target = objects[accessObject.Name];
            state = accessObject.CurrentView == AcCurrentView.acCurViewDesign
                ? DocumentState.DesignView
                : DocumentState.ActiveView;
            return new HostDocument(accessObject.Name, className, target, state);
        }
    }
}
