public abstract class TestsBase : IDisposable
{
    protected TestsBase()
    {
        SystemTime.Today = () => new DateTime(2023, 10, 10);
    }

    public void Dispose()
    {
        SystemTime.Reset();
    }
}