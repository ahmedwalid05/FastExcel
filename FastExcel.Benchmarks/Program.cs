using BenchmarkDotNet.Running;

namespace FastExcel.Benchmarks
{
    /// <summary>
    /// Entry point. Run everything with:
    ///     dotnet run -c Release --project FastExcel.Benchmarks -- --filter *
    /// or a single suite with:
    ///     dotnet run -c Release --project FastExcel.Benchmarks -- --filter *ReadBenchmarks*
    /// </summary>
    public static class Program
    {
        public static void Main(string[] args) =>
            BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
    }
}
