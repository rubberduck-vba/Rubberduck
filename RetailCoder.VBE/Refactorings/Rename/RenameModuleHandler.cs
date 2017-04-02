using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameModuleHandler : RenameHandlerBase
    {
        public RenameModuleHandler(RenameModel model, IMessageBox messageBox)
            : base(model, messageBox) { }

        override public string ErrorMessage { get { return RubberduckUI.RenameDialog_ModuleRenameError; } }

        override public void Rename()
        {
            var component = Model.Target.QualifiedName.QualifiedModuleName.Component;
            var module = component.CodeModule;
            if (module.IsWrappingNullReference)
            {
                return;
            }

            RenameUsages(Model.Target);

            if (component.Type == ComponentType.Document)
            {
                var properties = component.Properties;
                var property = properties["_CodeName"];
                {
                    Rewrite();
                    property.Value = Model.NewName;
                }
            }
            else if (component.Type == ComponentType.UserForm)
            {
                var properties = component.Properties;
                var property = properties["Caption"];
                {
                    Rewrite();
                    if ((string)property.Value == Model.Target.IdentifierName)
                    {
                        property.Value = Model.NewName;
                    }
                    component.Name = Model.NewName;
                }
            }
            else
            {
                Rewrite();
                module.Name = Model.NewName;
            }
        }
    }
}
