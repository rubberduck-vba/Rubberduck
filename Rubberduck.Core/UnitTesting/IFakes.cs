namespace Rubberduck.UnitTesting
{
    public interface IFakes
    {
        void StartTest();
        void StopTest();
    }

    public interface IFakesFactory
    {
        IFakes Create();
        void Release(IFakes fakes);
    }
}
