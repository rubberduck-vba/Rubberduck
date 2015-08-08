using System.Drawing;
using Rubberduck.Properties;

namespace Rubberduck.UI.Command
{
    public class RefactorRemoveParametersCommand : ICommand
    {
        public void Execute()
        {
            throw new System.NotImplementedException();
        }
    }

    public class RefactorRemoveParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorRemoveParametersCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_RemoveParameter"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.RemoveParameters; } }
        public override Image Image { get { return Resources.RemoveParameters_6781_32; } }
        public override Image Mask { get { return Resources.RemoveParameters_6781_32_Mask; }}
    }
}