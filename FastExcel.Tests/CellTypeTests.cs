using System;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// How each kind of cell is understood on the way in.
    ///
    /// A cell declares its kind with a <c>t</c> attribute, and OOXML defines six values:
    /// <c>b</c> boolean, <c>e</c> error, <c>inlineStr</c> text stored in the cell itself,
    /// <c>n</c> number, <c>s</c> an index into the shared-string table, and <c>str</c> the cached
    /// string result of a formula. Omitting <c>t</c> means <c>n</c>.
    ///
    /// This library recognises exactly one of them. Everything except <c>t="s"</c> falls through
    /// to a branch that looks for a <c>&lt;v&gt;</c> child and, failing that, returns the
    /// concatenated text of the whole element. That is why an inline string comes back looking
    /// almost right — and why anything with more structure does not.
    ///
    /// The two that matter most in practice are <c>inlineStr</c>, which several report writers
    /// emit in preference to the shared-string table, and <c>b</c>, which anyone reading a sheet
    /// with a TRUE/FALSE column will hit immediately.
    /// </summary>
    public class CellTypeTests
    {
        private static object ReadFirstCell(string cellXml, params string[] sharedStrings)
        {
            using var workspace = new TempWorkspace();
            var builder = new XlsxBuilder();
            if (sharedStrings.Length > 0) builder = builder.WithSharedStrings(sharedStrings);

            var file = builder
                .WithSheet("Sheet1", XlsxBuilder.Row(1, cellXml))
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);
            return fastExcel.Read(1).Rows.Single().Cells.Single().Value;
        }

        // ------------------------------------------------------------------ the one that works

        [Fact]
        public void SharedString_ResolvesThroughTheTable()
        {
            Assert.Equal("shared", ReadFirstCell("<c r=\"A1\" t=\"s\"><v>0</v></c>", "shared"));
        }

        [Fact]
        public void NumberWithNoTypeAttribute_ReadsAsItsDigits()
        {
            // No `t` means numeric. The value arrives as a string either way - the library never
            // converts, so a caller casting to double has to parse it themselves.
            var value = ReadFirstCell("<c r=\"A1\"><v>42</v></c>");

            Assert.Equal("42", value);
            Assert.IsType<string>(value);
        }

        [Fact]
        public void NumberWithAnExplicitNumericType_ReadsTheSame()
        {
            Assert.Equal("42", ReadFirstCell("<c r=\"A1\" t=\"n\"><v>42</v></c>"));
        }

        [Fact]
        public void Decimal_KeepsItsInvariantPointOnTheWayIn()
        {
            Assert.Equal("1.5", ReadFirstCell("<c r=\"A1\"><v>1.5</v></c>"));
        }

        // ------------------------------------------------------------------ inline strings

        [Fact]
        public void InlineString_HappensToReadCorrectly()
        {
            // <is><t>text</t></is> has no <v>, so the fallback returns the element's concatenated
            // text - which for this shape is exactly the string. It works by coincidence rather
            // than by handling, which the next two tests show.
            Assert.Equal("inline text",
                ReadFirstCell("<c r=\"A1\" t=\"inlineStr\"><is><t>inline text</t></is></c>"));
        }

        [Fact]
        public void InlineStringWithFormattingRuns_IsConcatenatedIntoOneValue()
        {
            // Rich text splits a string into runs. Concatenating them is the right answer here,
            // and this is the one place the fallback gets rich text right - unlike the shared
            // string table, where the same shape shifts every later index (#10).
            Assert.Equal("Hello World",
                ReadFirstCell("<c r=\"A1\" t=\"inlineStr\"><is>" +
                              "<r><t>Hello </t></r><r><t>World</t></r>" +
                              "</is></c>"));
        }

        [Fact]
        public void InlineStringWithPhoneticHints_SwallowsTheHintIntoTheValue()
        {
            // <rPh> carries a pronunciation hint for East Asian text and is not part of the
            // value. Because the fallback takes all descendant text, the hint is appended to the
            // word - so a Japanese workbook reads back with its furigana glued on.
            var value = ReadFirstCell("<c r=\"A1\" t=\"inlineStr\"><is>" +
                                      "<t>東京</t><rPh sb=\"0\" eb=\"2\"><t>トウキョウ</t></rPh>" +
                                      "</is></c>");

            Assert.Equal("東京トウキョウ", value);

            KnownBug.StillBroken("#10",
                "a phonetic hint is not part of the cell's value, so a cell reading 東京 " +
                "returns just those two characters rather than the reading appended to them",
                () => Assert.Equal("東京", value));
        }

        [Fact]
        public void InlineStringWithNoTextAtAll_ReadsAsEmpty()
        {
            // A real shape: some writers emit t="inlineStr" with no <is> child.
            Assert.Equal(string.Empty, ReadFirstCell("<c r=\"A1\" t=\"inlineStr\"/>"));
        }

        // ------------------------------------------------------------------ booleans

        [Fact]
        public void Boolean_ReadsAsTheRawDigitRatherThanABoolean()
        {
            // OOXML stores booleans as 0 or 1 with t="b". The library returns the digit as a
            // string, so a caller cannot tell TRUE from the number 1, and casting to bool throws.
            var value = ReadFirstCell("<c r=\"A1\" t=\"b\"><v>1</v></c>");

            Assert.Equal("1", value);

            KnownBug.StillBroken("#77",
                "a cell marked t=\"b\" is returned as a bool, so a caller can cast it",
                () => Assert.Equal(true, value));
        }

        [Fact]
        public void BooleanFalse_IsIndistinguishableFromTheNumberZero()
        {
            Assert.Equal("0", ReadFirstCell("<c r=\"A1\" t=\"b\"><v>0</v></c>"));
        }

        // ------------------------------------------------------------------ errors

        [Theory]
        [InlineData("#DIV/0!")]
        [InlineData("#VALUE!")]
        [InlineData("#REF!")]
        [InlineData("#N/A")]
        [InlineData("#NAME?")]
        [InlineData("#NULL!")]
        [InlineData("#NUM!")]
        public void ErrorCell_ReadsAsTheErrorTextWithNothingMarkingItAsAnError(string error)
        {
            // t="e" says this cell holds a formula error, not data. It comes back as an ordinary
            // string, so a column of numbers containing one #DIV/0! looks like a column of
            // numbers containing one odd label.
            var value = ReadFirstCell("<c r=\"A1\" t=\"e\"><v>" + error + "</v></c>");

            Assert.Equal(error, value);
        }

        // ------------------------------------------------------------------ formulas

        [Fact]
        public void FormulaWithACachedNumber_ReturnsTheCachedValue()
        {
            // The important one: return the cached result, not the formula text. This passes -
            // it was fixed for #21 - so the test is a regression guard.
            Assert.Equal("3", ReadFirstCell("<c r=\"A1\"><f>1+2</f><v>3</v></c>"));
        }

        [Fact]
        public void FormulaWithACachedString_ReturnsTheCachedValue()
        {
            Assert.Equal("result", ReadFirstCell("<c r=\"A1\" t=\"str\"><f>A2&amp;A3</f><v>result</v></c>"));
        }

        [Fact]
        public void FormulaWithNoCachedValue_ReturnsTheFormulaText()
        {
            // Excel omits <v> when it has not calculated the sheet. With no cached value the
            // fallback returns the element's text, which is the formula itself - so the caller
            // gets "SUM(A2:A3)" where they expected a number.
            var value = ReadFirstCell("<c r=\"A1\"><f>SUM(A2:A3)</f></c>");

            Assert.Equal("SUM(A2:A3)", value);

            KnownBug.StillBroken("#68",
                "an uncalculated formula cell is distinguishable from a cell whose value is the " +
                "literal text of a formula",
                () => Assert.NotEqual("SUM(A2:A3)", value));
        }

        [Fact]
        public void SharedFormulaWithACachedValue_ReturnsTheCachedValue()
        {
            Assert.Equal("7",
                ReadFirstCell("<c r=\"A1\"><f t=\"shared\" si=\"0\" ref=\"A1:A3\">B1*2</f><v>7</v></c>"));
        }

        // ------------------------------------------------------------------ dates

        [Fact]
        public void ADate_ComesBackAsItsSerialNumber()
        {
            // There is no date type in xlsx. A date is a number plus a style whose number format
            // looks like a date, and the library never opens styles.xml - so every date arrives
            // as the raw serial. 45418 is 2024-05-06.
            var value = ReadFirstCell("<c r=\"A1\" s=\"1\"><v>45418</v></c>");

            Assert.Equal("45418", value);

            KnownBug.StillBroken("#58",
                "a cell whose number format is a date is returned as a DateTime rather than as " +
                "its underlying serial number",
                () => Assert.IsType<DateTime>(value));
        }

        [Fact]
        public void AStyleIndexPointingNowhere_DoesNotPreventTheCellBeingRead()
        {
            // Styles are ignored entirely, so an out-of-range style index cannot break a read.
            // That is worth pinning: it stops being true the moment dates are implemented.
            Assert.Equal("1", ReadFirstCell("<c r=\"A1\" s=\"9999\"><v>1</v></c>"));
        }

        // ------------------------------------------------------------------ empty and odd shapes

        [Fact]
        public void ACellWithNoValueElement_ReadsAsEmpty()
        {
            Assert.Equal(string.Empty, ReadFirstCell("<c r=\"A1\"/>"));
        }

        [Fact]
        public void ACellWithAnEmptyValueElement_ReadsAsEmpty()
        {
            Assert.Equal(string.Empty, ReadFirstCell("<c r=\"A1\"><v></v></c>"));
        }

        [Fact]
        public void ASharedStringCellWithAnEmptyIndex_ReadsAsEmpty()
        {
            // Emitted in the wild by AG Grid: t="s" with no index at all.
            Assert.Equal(string.Empty, ReadFirstCell("<c r=\"A1\" t=\"s\"><v></v></c>", "unused"));
        }

        [Fact]
        public void ACellWithTwoValueElements_Throws()
        {
            // Malformed, but worth pinning: the lookup uses SingleOrDefault, so a duplicate <v>
            // throws InvalidOperationException rather than picking one or reporting bad input.
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSheet("Sheet1", XlsxBuilder.Row(1, "<c r=\"A1\"><v>1</v><v>2</v></c>"))
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);
            var thrown = Record.Exception(() => fastExcel.Read(1).Rows.Single().Cells.ToList());

            Assert.IsType<InvalidOperationException>(thrown);
        }

        [Fact]
        public void AnUnknownTypeAttribute_FallsBackRatherThanFailing()
        {
            // A future or misspelled `t` value takes the same fallback as everything else.
            Assert.Equal("value", ReadFirstCell("<c r=\"A1\" t=\"d\"><v>value</v></c>"));
        }

        // ------------------------------------------------------------------ text fidelity

        [Theory]
        [InlineData("plain")]
        [InlineData("with spaces")]
        [InlineData("  leading")]
        [InlineData("trailing  ")]
        [InlineData("café")]
        [InlineData("你好")]
        [InlineData("\U0001F600")]
        [InlineData("_x0041_")]
        public void SharedStringText_SurvivesTheReadIntact(string text)
        {
            // The last case is the interesting one. "_x0041_" is a legitimate string that a user
            // could store, and it is also the escape sequence for "A". The loader decodes every
            // entry unconditionally, so a real value shaped like an escape is silently rewritten.
            var value = ReadFirstCell("<c r=\"A1\" t=\"s\"><v>0</v></c>", text);

            if (text == "_x0041_")
            {
                Assert.Equal("A", value);

                KnownBug.StillBroken("#76",
                    "a stored string that merely looks like an escape sequence is returned as " +
                    "written; today \"_x0041_\" is decoded into \"A\" on the way in",
                    () => Assert.Equal("_x0041_", value));
                return;
            }

            Assert.Equal(text, value);
        }

        [Fact]
        public void InlineStringText_KeepsItsWhitespaceWhenPreserveIsSet()
        {
            Assert.Equal("  padded  ",
                ReadFirstCell("<c r=\"A1\" t=\"inlineStr\"><is>" +
                              "<t xml:space=\"preserve\">  padded  </t></is></c>"));
        }
    }
}
