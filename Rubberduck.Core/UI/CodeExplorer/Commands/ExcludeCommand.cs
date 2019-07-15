//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Threading;
//using Rubberduck.Interaction;
//using Rubberduck.Navigation.CodeExplorer;
//using Rubberduck.Parsing.VBA;
//using Rubberduck.Resources.CodeExplorer;
//using Rubberduck.VBEditor.ComManagement;
//using Rubberduck.VBEditor.SafeComWrappers;
//using Rubberduck.VBEditor.SafeComWrappers.Abstract;

//namespace Rubberduck.UI.CodeExplorer.Commands
//{
//    public class ExcludeCommand : CodeExplorerCommandBase
//    {
//        private readonly IParseManager _parseManager;
//        private readonly IProjectsRepository _projectsRepository;
//        private readonly IMessageBox _messageBox;
//        private readonly IVBE _vbe;

//        private static readonly Type[] ApplicableNodes =
//        {
//            typeof(CodeExplorerComponentViewModel)
//        };


//        public ExcludeCommand(IParseManager parseManager, IProjectsRepository projectsRepository, IMessageBox messageBox, IVBE vbe)
//        {
//            _parseManager = parseManager;
//            _projectsRepository = projectsRepository;
//            _messageBox = messageBox;
//            _vbe = vbe;
//        }

//        public override IEnumerable<Type> ApplicableNodeTypes => ApplicableNodes;

//        protected override void OnExecute(object parameter)
//        {
//            if (!(parameter is CodeExplorerComponentViewModel node) ||
//                node.Declaration == null ||
//                node.Declaration.QualifiedName.QualifiedModuleName.ComponentType == ComponentType.Document)
//            {
//                return;
//            }
            
//            var qualifiedModuleName = node.Declaration.QualifiedName.QualifiedModuleName;           

//            try
//            {                
//                _projectsRepository.RemoveComponent(qualifiedModuleName);
//            }
//            catch (Exception ex)
//            {
//                _messageBox.NotifyWarn(ex.Message, string.Format(CodeExplorerUI.RemoveError_Caption, qualifiedModuleName.ComponentName)); // TODO
//            }
//        }
//    }
//}
