public static class SystemTime
{
    // Default behavior: returns the real DateTime.Today
    public static Func<DateTime> Today = () => DateTime.Today;

    // Method to reset Today to the default behavior
    public static void Reset()
    {
        Today = () => DateTime.Today;
    }
}