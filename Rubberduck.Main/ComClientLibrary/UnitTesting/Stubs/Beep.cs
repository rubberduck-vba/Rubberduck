using Rubberduck.Runtime;

namespace Rubberduck.UnitTesting
{
    internal class Beep : StubBase
    {
        public Beep(IBeepInterceptor interceptor)
        {
            interceptor.Beep += BeepCallback;
        }

        public void BeepCallback(object sender, BeepEventArgs e)
        {
            OnCallBack(true);

            if (PassThrough)
            {
                VbeProvider.VbeRuntime.Beep();
            }

            e.Handled = true;
        }
    }
}
