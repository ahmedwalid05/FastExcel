using System;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// Column letter/number conversion. These are pure functions on the public surface, they
    /// underpin every cell reference the library reads or writes, and they had no coverage at
    /// all — so they are exercised exhaustively across Excel's full column range here.
    /// </summary>
    public class ColumnNameTests
    {
        [Theory]
        [InlineData(1, "A")]
        [InlineData(2, "B")]
        [InlineData(25, "Y")]
        [InlineData(26, "Z")]
        [InlineData(27, "AA")]      // first two-letter column
        [InlineData(28, "AB")]
        [InlineData(51, "AY")]
        [InlineData(52, "AZ")]
        [InlineData(53, "BA")]
        [InlineData(702, "ZZ")]     // last two-letter column
        [InlineData(703, "AAA")]    // first three-letter column
        [InlineData(16384, "XFD")]  // Excel's maximum column
        public void NumberToName(int number, string expected)
        {
            Assert.Equal(expected, Cell.GetExcelColumnName(number));
        }

        [Theory]
        [InlineData("A", 1)]
        [InlineData("Z", 26)]
        [InlineData("AA", 27)]
        [InlineData("AZ", 52)]
        [InlineData("BA", 53)]
        [InlineData("ZZ", 702)]
        [InlineData("AAA", 703)]
        [InlineData("XFD", 16384)]
        public void NameToNumber(string name, int expected)
        {
            Assert.Equal(expected, Cell.GetExcelColumnNumber(name));
        }

        [Theory]
        [InlineData("A1", 1)]
        [InlineData("B12", 2)]
        [InlineData("AA100", 27)]
        [InlineData("XFD1048576", 16384)]
        public void NameToNumber_StripsTheRowNumber(string reference, int expected)
        {
            Assert.Equal(expected, Cell.GetExcelColumnNumber(reference));
        }

        [Fact]
        public void NumberToName_RoundTripsAcrossEveryValidColumn()
        {
            for (var number = 1; number <= 16384; number++)
            {
                var name = Cell.GetExcelColumnName(number);
                Assert.Equal(number, Cell.GetExcelColumnNumber(name, includesRowNumber: false));
            }
        }

        [Fact]
        public void NumberToName_NeverProducesAnEmptyOrNonAlphabeticName()
        {
            for (var number = 1; number <= 16384; number++)
            {
                var name = Cell.GetExcelColumnName(number);
                Assert.NotEqual(string.Empty, name);
                Assert.All(name, c => Assert.InRange(c, 'A', 'Z'));
            }
        }

        [Fact]
        public void Cell_RejectsAColumnNumberBelowOne()
        {
            // Column numbers are 1-based; 0 would silently produce an empty column name.
            Assert.Throws<Exception>(() => new Cell(0, "value"));
            Assert.Throws<Exception>(() => new Cell(-1, "value"));
        }

        [Fact]
        public void Row_RejectsARowNumberBelowOne()
        {
            Assert.Throws<Exception>(() => new Row(0, Array.Empty<Cell>()));
            Assert.Throws<Exception>(() => new Row(-1, Array.Empty<Cell>()));
        }
    }
}
