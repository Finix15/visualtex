using System.Text;
using System.Xml.Linq;

namespace VisualTeX.MathTypeConversion;

internal sealed class Mtef5MathMlConverter
{
    private static readonly XNamespace M = "http://www.w3.org/1998/Math/MathML";
    private readonly byte[] _data;
    private int _position;
    private int _records;
    private bool _display;
    private string _product = string.Empty;
    private int _sourceCharacterCount;
    private readonly List<string> _encodings = new() { string.Empty, "MTCode", "Unknown", "Symbol", "MTExtra" };
    private readonly List<int> _fontEncodings = new() { 0 };

    private Mtef5MathMlConverter(byte[] data) => _data = data;

    public static MathTypeDecodeResult Convert(byte[] data)
    {
        var reader = new Mtef5MathMlConverter(data);
        return reader.Read();
    }

    private MathTypeDecodeResult Read()
    {
        if (ReadByte() != 5) Fail("MTEF_VERSION_UNSUPPORTED", "Only MathType MTEF v5 is supported.", true);
        ReadByte(); // platform
        ReadByte(); // product
        var major = ReadByte();
        var minor = ReadByte();
        _product = $"MathType {major}.{minor}";
        ReadNullTerminatedAscii(64);
        _display = (ReadByte() & 1) == 0;

        XElement? body = null;
        while (_position < _data.Length && body is null)
        {
            CountRecord();
            var type = ReadByte();
            switch (type)
            {
                case 1: body = Row(ReadLine(1)); break;
                case 4: body = ReadPile(1); break;
                case 8: ReadFontStyleDefinition(); break;
                case 9: ReadSize(); break;
                case >= 10 and <= 14: break;
                case 15: ReadUnsigned(); break;
                case 16: ReadColorDefinition(); break;
                case 17: ReadFontDefinition(); break;
                case 18: ReadEquationPreferences(); break;
                case 19: ReadEncodingDefinition(); break;
                case >= 100: SkipFuture(); break;
                default: Fail("MTEF_ROOT_RECORD_INVALID", $"Unexpected MTEF record {type} before the equation body."); break;
            }
        }
        if (body is null) Fail("MTEF_BODY_MISSING", "The MTEF equation body is missing.");
        if (ReadByte() != 0) Fail("MTEF_FINAL_END_MISSING", "The MTEF equation is missing its final END record.");
        if (_position != _data.Length) Fail("MTEF_TRAILING_DATA", "The MTEF stream contains records after the final END record.");
        var math = new XElement(M + "math", new XAttribute("display", _display ? "block" : "inline"), body);
        var emitted = math.DescendantsAndSelf().Count(value => value.Attribute("data-mtef-source") is not null);
        if (emitted != _sourceCharacterCount)
            Fail("SOURCE_DESTINATION_MISMATCH", "Not every source character was represented in the MathML tree.", true);
        foreach (var marker in math.DescendantsAndSelf().Attributes("data-mtef-source").ToList()) marker.Remove();
        ValidateMathMl(math);
        return new MathTypeDecodeResult
        {
            Status = MathTypeParseStatus.Convertible,
            MathMl = math.ToString(SaveOptions.DisableFormatting), Display = _display,
            MtefVersion = 5, Product = _product,
        };
    }

    private List<XElement> ReadContainer(int depth, bool stopAtEnd)
    {
        EnsureDepth(depth);
        var nodes = new List<XElement>();
        while (_position < _data.Length)
        {
            CountRecord();
            var type = ReadByte();
            if (type == 0)
            {
                if (stopAtEnd) return nodes;
                continue;
            }
            switch (type)
            {
                case 1: nodes.Add(Row(ReadLine(depth + 1))); break;
                case 2: nodes.Add(ReadCharacter(depth + 1)); break;
                case 3: nodes.Add(ReadTemplate(nodes, depth + 1)); break;
                case 4: nodes.Add(ReadPile(depth + 1)); break;
                case 5: nodes.Add(ReadMatrix(depth + 1)); break;
                case 6: SkipEmbellishment(); break;
                case 7: SkipRuler(); break;
                case 8: ReadFontStyleDefinition(); break;
                case 9: ReadSize(); break;
                case >= 10 and <= 14: break;
                case 15: ReadUnsigned(); break;
                case 16: ReadColorDefinition(); break;
                case 17: ReadFontDefinition(); break;
                case 18: ReadEquationPreferences(); break;
                case 19: ReadEncodingDefinition(); break;
                case >= 100: SkipFuture(); break;
                default: throw new InvalidDataException($"Unsupported MTEF record type {type} in equation body.");
            }
        }
        if (stopAtEnd) Fail("MTEF_LIST_UNTERMINATED", "The MTEF object list is unterminated.");
        return nodes;
    }

    private List<XElement> ReadLine(int depth)
    {
        var options = ReadByte();
        SkipNudge(options);
        if ((options & 0x04) != 0) Skip(2);
        if ((options & 0x02) != 0)
        {
            if (ReadByte() != 7) throw new InvalidDataException("The MTEF line ruler is invalid.");
            SkipRuler();
        }
        if ((options & 0x01) != 0) return new List<XElement>();
        return ReadContainer(depth, true);
    }

    private XElement ReadCharacter(int depth)
    {
        var options = ReadByte();
        SkipNudge(options);
        var typeface = ReadSigned();
        int code = -1;
        if ((options & 0x20) == 0) code = ReadUInt16();
        if ((options & 0x04) != 0) ReadByte();
        if ((options & 0x10) != 0) ReadUInt16();
        if ((options & 0x20) != 0)
            Fail("ENCODING_UNSUPPORTED", "A character without an MTCode value cannot be mapped safely.", true);
        if (code < 0 || code > 0x10ffff || code is >= 0xd800 and <= 0xdfff || code is >= 0xe000 and <= 0xf8ff)
            Fail("MTCODE_UNMAPPED", "The MTEF character code cannot be mapped safely.", true);
        var value = char.ConvertFromUtf32(code);
        _sourceCharacterCount++;
        XElement node;
        if (string.IsNullOrWhiteSpace(value) || typeface == 24) node = new XElement(M + "mspace", new XAttribute("width", "0.25em"));
        else if (char.IsDigit(value, 0)) node = new XElement(M + "mn", value);
        else if (IsOperator(value)) node = new XElement(M + "mo", value);
        else node = new XElement(M + "mi", value);
        node.SetAttributeValue("data-mtef-source", "1");
        if ((options & 0x02) != 0 || typeface == 2)
            node.SetAttributeValue("mathvariant", "normal");
        if ((options & 0x01) != 0)
            node = ApplyEmbellishments(node, depth);
        return node;
    }

    private XElement ReadTemplate(List<XElement> preceding, int depth)
    {
        var options = ReadByte();
        SkipNudge(options);
        var selector = ReadByte();
        var variation = ReadVariation();
        ReadByte(); // template options
        var children = ReadContainer(depth, true);
        var slots = children.Where(child => child.Name == M + "mrow").ToList();
        var structuralCharacters = children.Where(child => child.Name != M + "mrow").ToList();
        foreach (var structural in structuralCharacters)
        {
            if (!IsExpectedStructuralCharacter(selector, structural.Value))
                Fail("TEMPLATE_CHARACTER_UNEXPECTED", $"Template {selector} contains an unexpected structural character.", true);
            var marker = structural.DescendantsAndSelf().Attributes("data-mtef-source").SingleOrDefault();
            if (marker is null) Fail("TEMPLATE_RECORD_UNEXPECTED", "A template contains an unexpected non-slot record.", true);
            marker!.Remove();
            _sourceCharacterCount--;
        }
        if (slots.Count == 0) slots.Add(Row(children));
        XElement Slot(int index) => index < slots.Count ? slots[index] : new XElement(M + "mrow");

        if (selector == 26) throw new NotSupportedException("MathType long division is not supported by Word OMML conversion.");
        if (selector is 21 or 22) Fail("CUSTOM_BIG_OPERATOR_UNSUPPORTED", "A custom MathType big operator cannot be mapped safely.", true);
        if (selector is >= 27 and <= 29)
        {
            if ((variation & 1) != 0) Fail("PRECEDING_SCRIPT_UNSUPPORTED", "A script preceding its base is not supported safely.", true);
            if (preceding.Count == 0) Fail("SCRIPT_BASE_MISSING", "A MathType script has no base expression.", true);
            var @base = preceding[preceding.Count - 1];
            preceding.RemoveAt(preceding.Count - 1);
            var subscript = Slot(0);
            var superscript = Slot(1);
            return selector switch
            {
                27 when !IsEmpty(subscript) => new XElement(M + "msub", @base, subscript),
                28 when !IsEmpty(superscript) => new XElement(M + "msup", @base, superscript),
                29 when !IsEmpty(subscript) && !IsEmpty(superscript) => new XElement(M + "msubsup", @base, subscript, superscript),
                _ => throw new MtefParseException("SCRIPT_SLOT_EMPTY", "A required MathType script slot is empty.", _position, true),
            };
        }
        return selector switch
        {
            <= 9 => Fence(selector, variation, Slot(0)),
            10 => (variation & 1) != 0
                ? new XElement(M + "mroot", Slot(Math.Min(1, slots.Count - 1)), Slot(0))
                : new XElement(M + "msqrt", Slot(slots.Count - 1)),
            11 => new XElement(M + "mfrac", Slot(0), Slot(1)),
            12 => Accent("munder", Slot(0), "_"),
            13 => Accent("mover", Slot(0), "¯"),
            14 => Accent((variation & 0x08) != 0 ? "munder" : "mover", Slot(0), (variation & 0x10) != 0 ? "←" : "→"),
            >= 15 and <= 20 => BigOperator(selector, variation, slots, structuralCharacters.SingleOrDefault()?.Value),
            23 => Limit(slots),
            24 or 25 => Accent((variation & 1) != 0 ? "mover" : "munder", Slot(0), selector == 24 ? "⏞" : "⎴"),
            30 => new XElement(M + "mrow", new XElement(M + "mo", "⟨"), Slot(0), new XElement(M + "mo", "|"), Slot(1), new XElement(M + "mo", "⟩")),
            31 => Accent((variation & 4) != 0 ? "munder" : "mover", Slot(0), (variation & 1) != 0 ? "←" : "→"),
            32 => Accent("mover", Slot(0), "~"),
            33 => Accent("mover", Slot(0), "^"),
            34 => Accent("mover", Slot(0), "⌒"),
            35 => Accent("mover", Slot(0), "⌢"),
            36 => new XElement(M + "menclose", new XAttribute("notation", "updiagonalstrike"), Slot(0)),
            37 => new XElement(M + "menclose", new XAttribute("notation", "box"), Slot(0)),
            _ => throw new NotSupportedException($"MathType template {selector} is not supported."),
        };
    }

    private XElement ReadPile(int depth)
    {
        var options = ReadByte();
        SkipNudge(options);
        ReadByte(); ReadByte();
        if ((options & 0x02) != 0) { if (ReadByte() != 7) throw new InvalidDataException("Invalid pile ruler."); SkipRuler(); }
        var children = ReadContainer(depth, true);
        return new XElement(M + "mtable", children.Select(child => new XElement(M + "mtr", new XElement(M + "mtd", child))));
    }

    private XElement ReadMatrix(int depth)
    {
        var options = ReadByte();
        SkipNudge(options);
        ReadByte(); ReadByte(); ReadByte();
        var rows = ReadByte();
        var columns = ReadByte();
        if (rows == 0 || columns == 0 || rows * columns > 4096) throw new InvalidDataException("The MTEF matrix dimensions are invalid.");
        Skip((rows + 4) / 4);
        Skip((columns + 4) / 4);
        var children = ReadContainer(depth, true);
        var slots = children.Where(child => child.Name == M + "mrow").ToList();
        if (slots.Count != rows * columns) Fail("MATRIX_SLOT_COUNT", "The MathType matrix cell count does not match its dimensions.");
        var table = new XElement(M + "mtable");
        for (var row = 0; row < rows; row++)
        {
            var tr = new XElement(M + "mtr");
            for (var column = 0; column < columns; column++)
                tr.Add(new XElement(M + "mtd", row * columns + column < slots.Count ? slots[row * columns + column] : new XElement(M + "mrow")));
            table.Add(tr);
        }
        return table;
    }

    private XElement ApplyEmbellishments(XElement node, int depth)
    {
        while (_position < _data.Length && _data[_position] != 0)
        {
            if (ReadByte() != 6) throw new InvalidDataException("The MTEF embellishment list is invalid.");
            var options = ReadByte();
            SkipNudge(options);
            var kind = ReadByte();
            node = kind switch
            {
                2 => Accent("mover", node, "˙"), 3 => Accent("mover", node, "¨"),
                4 => Accent("mover", node, "⃛"), 5 => Accent("mover", node, "′"),
                6 => Accent("mover", node, "″"), 7 => Accent("mover", node, "‴"),
                8 => Accent("mover", node, "′"), 9 => Accent("mover", node, "~"),
                10 => Accent("mover", node, "^") , 17 => Accent("mover", node, "¯"),
                20 => Accent("mover", node, "→"), 21 => Accent("mover", node, "←"),
                29 => Accent("munder", node, "_") , 30 => Accent("munder", node, "~"),
                33 => Accent("munder", node, "→"), 34 => Accent("munder", node, "←"),
                _ => throw new NotSupportedException($"MathType embellishment {kind} is not supported."),
            };
        }
        if (ReadByte() != 0) throw new InvalidDataException("The MTEF embellishment list is unterminated.");
        return node;
    }

    private static XElement Fence(int selector, int variation, XElement content)
    {
        var pair = selector switch
        {
            0 => ("⟨", "⟩"), 1 => ("(", ")"), 2 => ("{", "}"), 3 => ("[", "]"),
            4 => ("|", "|"), 5 => ("‖", "‖"), 6 => ("⌊", "⌋"), 7 => ("⌈", "⌉"),
            _ => ("(", ")"),
        };
        return new XElement(M + "mrow",
            (variation & 1) != 0 ? new XElement(M + "mo", new XAttribute("stretchy", "true"), pair.Item1) : null,
            content,
            (variation & 2) != 0 ? new XElement(M + "mo", new XAttribute("stretchy", "true"), pair.Item2) : null);
    }

    private XElement BigOperator(int selector, int variation, List<XElement> slots, string? sourceSymbol)
    {
        var symbol = selector switch { 15 => "∫", 16 => "∑", 17 => "∏", 18 => "∐", 19 => "⋃", 20 => "⋂", _ => "∑" };
        if (!string.Equals(sourceSymbol, symbol, StringComparison.Ordinal))
            Fail("BIG_OPERATOR_MISMATCH", "The MathType big-operator character does not match its template selector.", true);
        XElement op = new XElement(M + "mo", new XAttribute("largeop", "true"), symbol);
        var hasLower = (variation & 1) != 0;
        var hasUpper = (variation & 2) != 0;
        var body = slots.Count > 0 ? slots[0] : new XElement(M + "mrow");
        if (hasLower && hasUpper) op = new XElement(M + "munderover", op, slots.ElementAtOrDefault(2) ?? new XElement(M + "mrow"), slots.ElementAtOrDefault(1) ?? new XElement(M + "mrow"));
        else if (hasLower) op = new XElement(M + "munder", op, slots.ElementAtOrDefault(1) ?? new XElement(M + "mrow"));
        else if (hasUpper) op = new XElement(M + "mover", op, slots.ElementAtOrDefault(1) ?? new XElement(M + "mrow"));
        return new XElement(M + "mrow", op, body);
    }

    private static XElement Limit(List<XElement> slots)
    {
        var main = slots.ElementAtOrDefault(0) ?? new XElement(M + "mrow");
        if (slots.Count >= 3) return new XElement(M + "munderover", main, slots[1], slots[2]);
        if (slots.Count >= 2) return new XElement(M + "munder", main, slots[1]);
        return main;
    }

    private static XElement Accent(string element, XElement body, string mark) =>
        new(M + element, body, new XElement(M + "mo", new XAttribute("stretchy", "true"), mark));

    private static XElement Row(IEnumerable<XElement> nodes) => new(M + "mrow", nodes);
    private static bool IsOperator(string value) => "+−-=/<>≤≥±×÷·,;:()[]{}|∫∑∏∞→←".IndexOf(value, StringComparison.Ordinal) >= 0;

    private void SkipEmbellishment() { var options = ReadByte(); SkipNudge(options); ReadByte(); }
    private void SkipRuler() { var count = ReadByte(); Skip(count * 3); }
    private void SkipFuture() { var length = ReadUnsigned(); Skip(length); }
    private void SkipNudge(int options) { if ((options & 0x08) == 0) return; var x = ReadByte(); var y = ReadByte(); if (x == 128 && y == 128) Skip(4); }
    private int ReadSigned() { var value = ReadByte(); return value == 0xff ? unchecked((short)ReadUInt16()) : value - 128; }
    private int ReadUnsigned() { var first = ReadByte(); return first != 0xff ? first : ReadUInt16(); }
    private int ReadVariation() { var first = ReadByte(); return (first & 0x80) == 0 ? first : (first & 0x7f) | (ReadByte() << 8); }
    private int ReadUInt16() { var low = ReadByte(); return low | (ReadByte() << 8); }
    private byte ReadByte() { if (_position >= _data.Length) throw new InvalidDataException("The MTEF stream is truncated."); return _data[_position++]; }
    private void Skip(int count) { if (count < 0 || _position + count > _data.Length) throw new InvalidDataException("The MTEF stream is truncated."); _position += count; }
    private string ReadNullTerminatedAscii(int maximum) { var start = _position; while (_position < _data.Length && _data[_position] != 0 && _position - start < maximum) _position++; if (_position >= _data.Length || _position - start >= maximum) throw new InvalidDataException("The MTEF application key is invalid."); var value = Encoding.ASCII.GetString(_data, start, _position - start); _position++; return value; }
    private static void EnsureDepth(int depth) { if (depth > MathTypeLimits.MaximumDepth) throw new InvalidDataException("The MTEF nesting limit was exceeded."); }

    private void ReadFontStyleDefinition() { var index = ReadUnsigned(); ReadByte(); if (index <= 0 || index >= _fontEncodings.Count) Fail("FONT_REFERENCE_INVALID", "A font style references an undefined font."); }
    private void ReadFontDefinition()
    {
        var encoding = ReadUnsigned();
        if (encoding <= 0 || encoding >= _encodings.Count) Fail("ENCODING_REFERENCE_INVALID", "A font references an undefined encoding.");
        ReadNullTerminatedAscii(512);
        _fontEncodings.Add(encoding);
    }
    private void ReadEncodingDefinition() { _encodings.Add(ReadNullTerminatedAscii(512)); }
    private void ReadColorDefinition()
    {
        var options = ReadByte();
        Skip((options & 1) != 0 ? 8 : 6);
        if ((options & 4) != 0) ReadNullTerminatedAscii(512);
    }
    private void ReadSize()
    {
        var marker = ReadByte();
        if (marker == 101) { ReadUInt16(); return; }
        if (marker == 100) { ReadByte(); ReadUInt16(); return; }
        if (marker > 7) Fail("SIZE_RECORD_INVALID", "The MTEF SIZE record is invalid.");
        ReadByte();
    }
    private void ReadEquationPreferences()
    {
        ReadByte();
        SkipDimensionArray();
        SkipDimensionArray();
        var count = ReadByte();
        for (var i = 0; i < count; i++) { var font = ReadUnsigned(); if (font != 0) ReadByte(); }
    }
    private void SkipDimensionArray()
    {
        var count = ReadByte();
        var completed = 0;
        var high = true;
        byte packed = 0;
        while (completed < count)
        {
            if (high) packed = ReadByte();
            var nibble = high ? (packed >> 4) & 0xf : packed & 0xf;
            high = !high;
            if (nibble == 0xf) completed++;
        }
    }
    private void CountRecord() { if (++_records > MathTypeLimits.MaximumRecords) Fail("RECORD_LIMIT", "The MTEF record limit was exceeded."); }
    private void Fail(string code, string message, bool unsupported = false) => throw new MtefParseException(code, message, Math.Max(0, _position - 1), unsupported);
    private static bool IsEmpty(XElement value) => value.Name == M + "mrow" && !value.Nodes().Any();
    private static bool IsExpectedStructuralCharacter(int selector, string value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        if (selector <= 9) return "⟨⟩(){}[]|‖⌊⌋⌈⌉".IndexOf(value, StringComparison.Ordinal) >= 0;
        if (selector == 10) return value == "√";
        if (selector is >= 15 and <= 22) return "∫∑∏∐⋃⋂".IndexOf(value, StringComparison.Ordinal) >= 0;
        if (selector == 30) return "⟨|⟩".IndexOf(value, StringComparison.Ordinal) >= 0;
        return selector is 14 or 24 or 25 or >= 31 and <= 35;
    }
    private void ValidateMathMl(XElement math)
    {
        foreach (var identifier in math.Descendants(M + "mi"))
            if (string.IsNullOrWhiteSpace(identifier.Value)) Fail("WHITESPACE_IDENTIFIER", "Whitespace cannot be emitted as a MathML identifier.");
        foreach (var element in math.Descendants().Where(value => value.Name == M + "mfrac" || value.Name == M + "msub" || value.Name == M + "msup" || value.Name == M + "msubsup" || value.Name == M + "mroot"))
            if (element.Elements().Any(IsEmpty)) Fail("MANDATORY_SLOT_EMPTY", "A required MathML slot is empty.", true);
        foreach (var element in math.Descendants(M + "msqrt"))
            if (!element.Nodes().Any()) Fail("MANDATORY_SLOT_EMPTY", "A required square-root slot is empty.", true);
    }
}
