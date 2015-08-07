using System.Drawing;
using Rubberduck.Properties;

namespace Rubberduck.UI.Command
{
    public partial class RefactorExtractMethodCommand : ICommand
    {
        public void Execute()
        {
            throw new System.NotImplementedException();
        }
    }

    public class RefactorExtractMethodCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractMethodCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return RubberduckUI.RefactorMenu_ExtractMethod; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ExtractMethod; } }
        public override Image Image { get { return Resources.ExtractMethod_6786_32; } }
        public override Image Mask { get { return Resources.ExtractMethod_6786_32_Mask; } }
    }
}