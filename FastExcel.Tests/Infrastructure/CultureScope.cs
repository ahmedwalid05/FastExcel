using System;
using System.Globalization;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// Runs a block under a specific culture and restores the previous one afterwards.
    /// <para>
    /// The xlsx format requires an invariant decimal point regardless of the machine's locale,
    /// so a library that formats numbers with the ambient culture writes files that Excel
    /// rejects outside en-US. That class of bug is invisible unless tests actually run under a
    /// comma-decimal culture, which is what this exists for.
    /// </para>
    /// </summary>
    public sealed class CultureScope : IDisposable
    {
        private readonly CultureInfo _previousCulture;
        private readonly CultureInfo _previousUiCulture;

        public CultureScope(string name)
        {
            _previousCulture = CultureInfo.CurrentCulture;
            _previousUiCulture = CultureInfo.CurrentUICulture;

            var culture = new CultureInfo(name);
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = culture;
        }

        public void Dispose()
        {
            CultureInfo.CurrentCulture = _previousCulture;
            CultureInfo.CurrentUICulture = _previousUiCulture;
        }

        /// <summary>
        /// Cultures worth exercising, and why each one is here:
        ///  - de-DE / ru-RU / fr-FR use a comma as the decimal separator
        ///  - fr-FR also uses a narrow no-break space as the group separator
        ///  - tr-TR has the dotless-i casing rule, which breaks naive ToUpper/ToLower comparisons
        ///  - ar-SA historically resolves to a non-Gregorian calendar for date formatting
        /// </summary>
        public static readonly string[] Problematic = { "de-DE", "ru-RU", "fr-FR", "tr-TR", "ar-SA" };
    }
}
