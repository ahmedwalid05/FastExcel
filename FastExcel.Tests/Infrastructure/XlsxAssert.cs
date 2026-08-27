using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// Checks that a produced file is a legal package, not merely one FastExcel can read back.
    ///
    /// A FastExcel -> FastExcel round trip proves almost nothing here, because the library
    /// escapes and unescapes symmetrically: a file can be unreadable to Excel and every other
    /// consumer while still round-tripping cleanly through its own reader. #76 survived five
    /// years for exactly that reason.
    ///
    /// Two gates, both cheap:
    ///   * <see cref="Package"/> parses [Content_Types].xml and the relationship graph and
    ///     throws if either is malformed. This is the closest thing in .NET to asking
    ///     "would Excel offer to repair this?".
    ///   * every *.xml part is then handed to <see cref="XDocument"/>, which catches unescaped
    ///     text, a truncated part, and bytes left over from a shorter rewrite.
    /// </summary>
    internal static class XlsxAssert
    {
        /// <summary>Content types every xlsx must declare, keyed by part.</summary>
        private static readonly Dictionary<string, string> RequiredParts = new Dictionary<string, string>
        {
            ["/xl/workbook.xml"] =
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            ["/xl/_rels/workbook.xml.rels"] =
                "application/vnd.openxmlformats-package.relationships+xml",
            ["/_rels/.rels"] =
                "application/vnd.openxmlformats-package.relationships+xml",
        };

        /// <summary>
        /// Asserts the file is a well-formed OOXML package whose XML parts all parse.
        /// Call this from every test that writes a file.
        /// </summary>
        internal static void IsValidPackage(FileInfo file)
        {
            file.Refresh();
            Assert.True(file.Exists, $"no file was produced at {file.FullName}");

            EveryXmlPartParses(file);

            Dictionary<string, string> contentTypes;
            try
            {
                using var package = Package.Open(file.FullName, FileMode.Open, FileAccess.Read);
                contentTypes = package.GetParts()
                    .ToDictionary(p => p.Uri.ToString(), p => p.ContentType);
            }
            catch (Exception ex)
            {
                throw new XlsxInvalidException(
                    $"the package itself is malformed, so Excel would offer to repair it: {ex.Message}", ex);
            }

            foreach (var required in RequiredParts)
            {
                Assert.True(contentTypes.ContainsKey(required.Key),
                    $"the package is missing {required.Key}");
                Assert.Equal(required.Value, contentTypes[required.Key]);
            }
        }

        /// <summary>
        /// Parses every *.xml and *.rels entry. Kept separate because it is the half that
        /// catches bad *text* — the package layer is happy with a part full of unescaped
        /// ampersands as long as the zip and the relationships line up.
        /// </summary>
        internal static void EveryXmlPartParses(FileInfo file)
        {
            using var zip = ZipFile.OpenRead(file.FullName);
            foreach (var entry in zip.Entries)
            {
                if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                    !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                using var stream = entry.Open();
                try
                {
                    XDocument.Load(stream);
                }
                catch (Exception ex)
                {
                    throw new XlsxInvalidException(
                        $"{entry.FullName} is not well-formed XML: {ex.Message}", ex);
                }
            }
        }

        /// <summary>Returns one entry's bytes as text, exactly as another reader would see them.</summary>
        internal static string RawPart(FileInfo file, string entryName)
        {
            using var zip = ZipFile.OpenRead(file.FullName);
            var entry = zip.GetEntry(entryName);
            Assert.True(entry != null, $"the package has no {entryName}");
            using var reader = new StreamReader(entry.Open());
            return reader.ReadToEnd();
        }
    }

    /// <summary>Distinguishes "the file is invalid" from an ordinary failed assertion.</summary>
    internal sealed class XlsxInvalidException : Exception
    {
        internal XlsxInvalidException(string message, Exception innerException)
            : base(message, innerException) { }
    }
}
