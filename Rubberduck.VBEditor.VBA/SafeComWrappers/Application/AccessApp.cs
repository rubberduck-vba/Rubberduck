using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Access;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class AccessApp : HostApplicationBase<Microsoft.Office.Interop.Access.Application>
    {
        private const string FormNamePrefix = "Form_";
        private const string FormClassName = "Access.Form";
        private const string ReportNamePrefix = "Report_";
        private const string ReportClassName = "Access.Report";

        private readonly Lazy<IVBProject> _dbcProject;
        
        public AccessApp(IVBE vbe) : base(vbe, "Access", true)
        {
            _dbcProject = new Lazy<IVBProject>(() =>
            {
                using (var wizHook = new SafeIDispatchWrapper<WizHook>(Application.WizHook))
                {
                    wizHook.Target.Key = 51488399;
                    return new VBProject(wizHook.Target.DbcVbProject);
                }
            });
        }

        public override HostDocument GetDocument(QualifiedModuleName moduleName)
        {
            try
            {
                if (moduleName.ComponentName.StartsWith(FormNamePrefix))
                {
                    var name = moduleName.ComponentName.Substring(FormNamePrefix.Length);
                    using (var currentProject = new SafeIDispatchWrapper<_CurrentProject>(Application.CurrentProject))
                    using (var allForms = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllForms))
                    using (var accessObject = new SafeIDispatchWrapper<AccessObject>(allForms.Target[name]))
                    {
                        return LoadHostDocument(moduleName, FormClassName, accessObject);
                    }
                }

                if (moduleName.ComponentName.StartsWith(ReportNamePrefix))
                {
                    var name = moduleName.ComponentName.Substring(ReportNamePrefix.Length);
                    using (var currentProject = new SafeIDispatchWrapper<_CurrentProject>(Application.CurrentProject))
                    using (var allReports = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllReports))
                    using (var accessObject = new SafeIDispatchWrapper<AccessObject>(allReports.Target[name]))
                    {
                        return LoadHostDocument(moduleName, name, accessObject);
                    }
                }
            }
            catch(Exception ex)
            {
                //Log and ignore
                _logger.Log(LogLevel.Info, ex, $"Failed to get host document {moduleName.ToString()}");
            }

            return null;
        }

        public override IEnumerable<HostDocument> GetDocuments()
        {
            var list = new List<HostDocument>();

            foreach (var document in DocumentComponents())
            {
                var moduleName = new QualifiedModuleName(document);
                var name = string.Empty;
                var className = string.Empty;
                if (document.Name.StartsWith(FormNamePrefix))
                {
                    className = FormClassName;
                    name = document.Name.Substring(FormNamePrefix.Length);
                }
                else if(document.Name.StartsWith(ReportNamePrefix))
                {
                    className = ReportClassName;
                    name = document.Name.Substring(ReportNamePrefix.Length);
                }

                using (var project = document.ParentProject)
                {
                    var state = GetDocumentState(project, name, className);
                    list.Add(new HostDocument(moduleName, name, className, state, null));
                }
            }

            return list;
        }

        public override bool CanOpenDocumentDesigner(QualifiedModuleName moduleName)
        {
            return GetDocument(moduleName) != null;
        }

        public override bool TryOpenDocumentDesigner(QualifiedModuleName moduleName)
        {
            try
            {
                if (moduleName.ComponentName.StartsWith(FormNamePrefix))
                {
                    var name = moduleName.ComponentName.Substring(FormNamePrefix.Length);
                    using (var currentProject = new SafeIDispatchWrapper<_CurrentProject>(Application.CurrentProject))
                    using (var allForms = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllForms))
                    using (var accessObject = new SafeIDispatchWrapper<AccessObject>(allForms.Target[name]))
                    using (var doCmd = new SafeIDispatchWrapper<DoCmd>(Application.DoCmd))
                    {
                        if (accessObject.Target.IsLoaded &&
                            accessObject.Target.CurrentView != AcCurrentView.acCurViewDesign)
                        {
                            doCmd.Target.Close(AcObjectType.acForm, name);
                        }

                        if (!accessObject.Target.IsLoaded)
                        {
                            doCmd.Target.OpenForm(name, AcFormView.acDesign);
                        }

                        return accessObject.Target.IsLoaded &&
                               accessObject.Target.CurrentView == AcCurrentView.acCurViewDesign;
                    }
                }

                if (moduleName.ComponentName.StartsWith(ReportNamePrefix))
                {
                    var name = moduleName.ComponentName.Substring(ReportNamePrefix.Length);
                    using (var currentProject = new SafeIDispatchWrapper<_CurrentProject>(Application.CurrentProject))
                    using (var allReports = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllReports))
                    using (var accessObject = new SafeIDispatchWrapper<AccessObject>(allReports.Target[name]))
                    using (var doCmd = new SafeIDispatchWrapper<DoCmd>(Application.DoCmd))
                    {
                        if (accessObject.Target.IsLoaded &&
                            accessObject.Target.CurrentView != AcCurrentView.acCurViewDesign)
                        {
                            doCmd.Target.Close(AcObjectType.acReport, name);
                        }

                        if (!accessObject.Target.IsLoaded)
                        {
                            doCmd.Target.OpenReport(name, AcView.acViewDesign);
                        }

                        return accessObject.Target.IsLoaded &&
                               accessObject.Target.CurrentView == AcCurrentView.acCurViewDesign;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Log(LogLevel.Info, ex, $"Unable to open the document in design view for {moduleName.ToString()}");
            }

            return false;
        }

        private DocumentState GetDocumentState(IVBProject project, string name, string className)
        {
            if (!project.Equals(_dbcProject.Value))
            {
                return DocumentState.Inaccessible;
            }

            using (var currentProject = new SafeIDispatchWrapper<_CurrentProject>(Application.CurrentProject))
            {
                switch (className)
                {
                    case FormClassName:
                        using (var allForms = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllForms))
                        using (var accessObject = new SafeIDispatchWrapper<AccessObject>(allForms.Target[name]))
                        {
                            if (accessObject.Target.IsLoaded)
                            {
                                return DetermineDocumentState(accessObject.Target.CurrentView);
                            }
                        }

                        break;
                    case ReportClassName:
                        using (var allReports = new SafeIDispatchWrapper<AllObjects>(currentProject.Target.AllReports))
                        using (var accessObject = new SafeIDispatchWrapper<AccessObject>(allReports.Target[name]))
                        {
                            if (accessObject.Target.IsLoaded)
                            {
                                return DetermineDocumentState(accessObject.Target.CurrentView);
                            }
                        }

                        break;
                }
            }

            return DocumentState.Inaccessible;
        }

        private HostDocument LoadHostDocument(QualifiedModuleName moduleName, string className, SafeIDispatchWrapper<AccessObject> accessObject)
        {
            var state = DocumentState.Inaccessible;
            if (!accessObject.Target.IsLoaded)
            {
                return new HostDocument(moduleName, accessObject.Target.Name, className, state, null);
            }

            if (moduleName.ProjectName == _dbcProject.Value.Name && moduleName.ProjectId == _dbcProject.Value.HelpFile)
            {
                state = DetermineDocumentState(accessObject.Target.CurrentView);
            }

            return new HostDocument(moduleName, accessObject.Target.Name, className, state, null);
        }

        private static DocumentState DetermineDocumentState(AcCurrentView CurrentView)
        {
            return CurrentView == AcCurrentView.acCurViewDesign
                ? DocumentState.DesignView
                : DocumentState.ActiveView;
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
                if (_dbcProject.IsValueCreated)
                {
                    _dbcProject.Value.Dispose();
                }
            }

            base.Dispose(disposing);
        }
    }
}

