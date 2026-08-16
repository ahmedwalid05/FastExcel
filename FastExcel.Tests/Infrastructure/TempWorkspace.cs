using System;
using System.IO;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// An isolated temp directory for one test, removed on dispose.
    /// <para>
    /// The pre-existing tests write their output into the shared ResourcesTests folder under
    /// fixed names like temp.xlsx, which couples tests to each other and leaves artifacts
    /// behind. New tests get their own directory instead, so they can run in any order.
    /// </para>
    /// </summary>
    public sealed class TempWorkspace : IDisposable
    {
        private readonly string _root;

        public TempWorkspace()
        {
            _root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "fastexcel-tests", Guid.NewGuid().ToString("n"));
            Directory.CreateDirectory(_root);
        }

        /// <summary>A path inside this workspace. The file is not created.</summary>
        public string Path(string fileName) => System.IO.Path.Combine(_root, fileName);

        /// <summary>A path inside this workspace that is guaranteed not to exist yet.</summary>
        public FileInfo NewFile(string fileName = "out.xlsx")
        {
            var file = new FileInfo(Path(fileName));
            if (file.Exists) file.Delete();
            file.Refresh();
            return file;
        }

        /// <summary>Copies one of the checked-in fixtures into this workspace so it can be mutated safely.</summary>
        public FileInfo CopyFixture(string fixtureName, string asName = null)
        {
            var source = new FileInfo(System.IO.Path.Combine(TestFixtures.Directory, fixtureName));
            return source.CopyTo(Path(asName ?? fixtureName), overwrite: true);
        }

        public void Dispose()
        {
            try { if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true); }
            catch (IOException) { /* a leaked handle should never fail a test */ }
        }
    }

    /// <summary>Locations of the .xlsx fixtures copied next to the test assembly.</summary>
    public static class TestFixtures
    {
        public static string Directory => System.IO.Path.Combine(Environment.CurrentDirectory, "ResourcesTests");

        public static FileInfo Get(string name) => new FileInfo(System.IO.Path.Combine(Directory, name));

        /// <summary>An empty single-sheet workbook, usable as a write template.</summary>
        public static FileInfo Template => Get("template.xlsx");

        /// <summary>Contains a shared string table with a duplicate value and rich-text runs.</summary>
        public static FileInfo SameKey => Get("SameKey.xlsx");
    }
}
