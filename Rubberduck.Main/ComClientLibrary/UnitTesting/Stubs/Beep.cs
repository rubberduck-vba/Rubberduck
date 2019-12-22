using Rubberduck.Runtime;

namespace Rubberduck.UnitTesting
{
    internal class Beep : StubBase
    {
        private readonly IBeepInterceptor _interceptor;

        public Beep(IBeepInterceptor interceptor)
        {
            _interceptor = interceptor;
            _interceptor.Beep += BeepCallback;
        }

        public void BeepCallback(object sender, BeepEventArgs e)
        {
            OnCallBack(true);

            if (PassThrough)
            {
                VbeProvider.VbeNativeApi.Beep();
            }

            e.Handled = true;
        }

        public override void Dispose()
        {
            _interceptor.Beep -= BeepCallback;
            base.Dispose();
        }
    }
}
