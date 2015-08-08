using System.Drawing;
using Rubberduck.Properties;

namespace Rubberduck.UI.Command
{
    public class RefactorReorderParametersCommand : ICommand
    {
        public void Execute()
        {
            throw new System.NotImplementedException();
        }
    }

    public class RefactorReorderParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorReorderParametersCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_ReorderParameters"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ReorderParameters; } }
        public override Image Image { get { return Resources.ReorderParameters_6780_32; } }
        public override Image Mask { get { return Resources.ReorderParameters_6780_32_Mask; } }
    }
}