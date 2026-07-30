using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlConverter
{
    private const string MathNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string WordNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string NaryCharacters = "∑∏∐∫∬∭∮∯∰∱∲∳⨑⋀⋁⋂⋃";
    private const string ExtendedIntegralCharacters = "∯∰∱∲∳⨑";
    private static readonly object TransformLock = new();
    private static XslCompiledTransform? _mathMlToOmml;
    private static XslCompiledTransform? _ommlToMathMl;

    internal static Range Insert(
        Application application,
        Document targetDocument,
        Range insertionRange,
        string mathMl,
        bool display,
        bool includeLeadingTab = false,
        bool replaceTarget = false)
    {
        var omml = TransformMathMlToOmml(mathMl);
        var tempPath = CreateTemporaryDocx(
            omml,
            includeLeadingTab: display && includeLeadingTab,
            forceInline: false);
        Document? sourceDocument = null;
        OMaths? sourceMaths = null;
        OMath? sourceMath = null;
        Range? sourceRange = null;
        OMath? insertedMath = null;
        Range? result = null;
        try
        {
            sourceDocument = application.Documents.Open(
                FileName: tempPath,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            sourceMaths = sourceDocument.OMaths;
            if (sourceMaths.Count != 1)
                throw new InvalidDataException(
                    "The temporary OMML document did not contain exactly one equation.");
            sourceMath = sourceMaths[1];
            if (display && includeLeadingTab)
            {
                var paragraph = sourceMath.Range.Paragraphs[1];
                try
                {
                    sourceRange = paragraph.Range.Duplicate;
                    sourceRange.End = Math.Max(sourceRange.Start, sourceRange.End - 1);
                }
                finally { Release(paragraph); }
            }
            else
            {
                sourceRange = sourceMath.Range;
            }

            if (!replaceTarget)
                insertionRange.Collapse(WdCollapseDirection.wdCollapseStart);
            var insertionStart = insertionRange.Start;
            insertionRange.FormattedText = sourceRange.FormattedText;
            insertedMath = FindMathAtPosition(
                    targetDocument,
                    insertionStart)
                ?? throw new InvalidOperationException(
                    "Word did not materialize the inserted OMML equation.");
            insertedMath.Type = display
                ? WdOMathType.wdOMathDisplay
                : WdOMathType.wdOMathInline;
            insertedMath.BuildUp();
            result = insertedMath.Range.Duplicate;
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(insertedMath);
            Release(sourceRange);
            Release(sourceMath);
            Release(sourceMaths);
            if (sourceDocument is not null)
            {
                try { sourceDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(sourceDocument);
            try { File.Delete(tempPath); } catch { }
        }
    }

    internal static string TransformMathMlToOmml(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("VisualTeX did not provide MathML for the Word OMML formula.");
        mathMl = NormalizeNestedEmptyBaseScripts(mathMl);
        mathMl = NormalizeMathMlAccents(mathMl);
        mathMl = NormalizeNaryArguments(mathMl);
        var placeholderResult = ReplaceExtendedIntegralsWithOfficePlaceholders(mathMl);
        mathMl = placeholderResult.MathMl;
        var display = IsBlockMathMl(mathMl);
        var transform = GetTransform();
        var inputSettings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = 4_000_000,
        };
        var outputSettings = transform.OutputSettings?.Clone() ?? new XmlWriterSettings();
        outputSettings.OmitXmlDeclaration = true;
        outputSettings.Encoding = new UTF8Encoding(false);
        using var sourceText = new StringReader(mathMl);
        using var source = XmlReader.Create(sourceText, inputSettings);
        using var outputText = new StringWriter();
        using (var output = XmlWriter.Create(outputText, outputSettings))
            transform.Transform(source, output);
        var transformed = outputText.ToString();
        var omml = ExtractSingleOMath(transformed);
        omml = RestoreExtendedIntegralCharacters(omml, placeholderResult.NaryCharacters);
        omml = NormalizeExplicitUprightRuns(omml, mathMl);
        omml = NormalizeExplicitTableColumnAlignment(omml, mathMl);
        return NormalizeDisplayNaryOmml(omml, display);
    }

    internal static string NormalizeExplicitUprightRuns(string omml, string mathMl)
    {
        if (string.IsNullOrWhiteSpace(omml) || string.IsNullOrWhiteSpace(mathMl))
            return omml;

        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var mathMlDocument = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var uprightTokens = mathMlDocument
            .Descendants()
            .Where(element =>
            {
                if (element.Name.Namespace != presentationMath) return false;
                var variant = element.Attribute("mathvariant")?.Value ?? string.Empty;
                return variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) >= 0
                    || variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) >= 0;
            })
            .Select(element => element.Value)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToHashSet(StringComparer.Ordinal);
        if (uprightTokens.Count == 0) return omml;

        var ommlDocument = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var run in ommlDocument.Descendants(math + "r"))
        {
            var text = string.Concat(run.Elements(math + "t").Select(element => element.Value));
            if (!uprightTokens.Contains(text)) continue;

            var properties = run.Element(math + "rPr");
            var plainStyle = properties?.Element(math + "sty");
            if (plainStyle?.Attribute(math + "val")?.Value != "p") continue;

            plainStyle.Remove();
            if (properties!.Element(math + "nor") is null)
                properties.AddFirst(new XElement(math + "nor"));
        }

        return ommlDocument.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static string NormalizeMathMlAccents(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return mathMl;

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        var accentCharacters = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["^"] = "\u0302",
            ["~"] = "\u0303",
            ["→"] = "\u20D7",
            ["←"] = "\u20D6",
            ["˙"] = "\u0307",
            ["¨"] = "\u0308",
            ["ˇ"] = "\u030C",
            ["˘"] = "\u0306",
            ["´"] = "\u0301",
            ["`"] = "\u0300",
            ["˚"] = "\u030A",
        };

        foreach (var mover in document.Descendants(presentationMath + "mover").ToList())
        {
            var children = mover.Elements().ToArray();
            if (children.Length != 2) continue;
            var mark = children[1];
            if (mark.Name != presentationMath + "mo") continue;
            if (string.Equals(
                    mark.Attribute("accent")?.Value,
                    "false",
                    StringComparison.OrdinalIgnoreCase)
                || string.Equals(
                    mover.Attribute("accent")?.Value,
                    "false",
                    StringComparison.OrdinalIgnoreCase))
                continue;

            var sourceCharacter = mark.Value;
            if (!accentCharacters.TryGetValue(sourceCharacter, out var combiningCharacter))
                continue;

            // Office's MML2OMML.XSL only creates a native m:acc node when the
            // MathML mover is explicitly marked as an accent and the mark is a
            // combining accent character. MathJax emits spacing characters
            // such as ^, ˙ and → without accent=true, which Office otherwise
            // converts into m:limUpp or replacement glyphs/placeholder boxes.
            mover.SetAttributeValue("accent", "true");
            mark.SetAttributeValue("accent", "true");
            mark.SetAttributeValue("stretchy", null);
            mark.SetAttributeValue("data-mjx-pseudoscript", null);
            mark.Value = combiningCharacter;
        }

        return document.ToString(SaveOptions.DisableFormatting);
    }

    internal static string NormalizeNestedEmptyBaseScripts(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return mathMl;
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace mathMlNamespace = "http://www.w3.org/1998/Math/MathML";
        var simpleScriptNames = new HashSet<XName>
        {
            mathMlNamespace + "msub",
            mathMlNamespace + "msup",
        };
        var allScriptNames = new HashSet<XName>(simpleScriptNames)
        {
            mathMlNamespace + "msubsup",
        };
        var transparentWrappers = new HashSet<XName>
        {
            mathMlNamespace + "mrow",
            mathMlNamespace + "mstyle",
            mathMlNamespace + "mpadded",
            mathMlNamespace + "mphantom",
            mathMlNamespace + "semantics",
        };

        bool IsEmptyMathNode(XElement element)
        {
            if (element.Name == mathMlNamespace + "mspace") return false;
            if (element.Nodes().OfType<XText>().Any(node => !string.IsNullOrWhiteSpace(node.Value)))
                return false;
            var children = element.Elements().ToArray();
            return children.Length == 0 || children.All(IsEmptyMathNode);
        }

        bool IsOnlyContentOfOuterScriptArgument(XElement candidate)
        {
            XElement current = candidate;
            while (current.Parent is XElement parent)
            {
                if (transparentWrappers.Contains(parent.Name))
                {
                    if (parent.Elements().Any(sibling =>
                            sibling != current && !IsEmptyMathNode(sibling)))
                        return false;
                    if (parent.Nodes().OfType<XText>().Any(node =>
                            !string.IsNullOrWhiteSpace(node.Value)))
                        return false;
                    current = parent;
                    continue;
                }
                if (!allScriptNames.Contains(parent.Name)) return false;
                var children = parent.Elements().ToList();
                var position = children.IndexOf(current);
                return position >= 1;
            }
            return false;
        }

        foreach (var script in document
                     .Descendants()
                     .Where(element => simpleScriptNames.Contains(element.Name))
                     .Reverse()
                     .ToList())
        {
            var children = script.Elements().ToArray();
            if (children.Length < 2
                || !IsEmptyMathNode(children[0])
                || !IsOnlyContentOfOuterScriptArgument(script))
                continue;

            // MathJax represents sources such as f_{_{\\mathrm H}} as an
            // outer subscript whose argument contains another subscript with
            // an empty base. Office faithfully renders that empty base as a
            // dotted equation placeholder. Inside an existing script slot the
            // extra empty-base level carries no useful layout information, so
            // replace it with its visible script argument. Standalone empty-
            // base scripts are intentionally preserved for prescript/tensor
            // notation.
            script.ReplaceWith(new XElement(children[1]));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    internal static string NormalizeNaryArguments(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return mathMl;
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace mathMlNamespace = "http://www.w3.org/1998/Math/MathML";
        var limitNames = new HashSet<XName>
        {
            mathMlNamespace + "munder",
            mathMlNamespace + "mover",
            mathMlNamespace + "munderover",
            mathMlNamespace + "msub",
            mathMlNamespace + "msup",
            mathMlNamespace + "msubsup",
        };
        var display = string.Equals(
            document.Root?.Attribute("display")?.Value,
            "block",
            StringComparison.OrdinalIgnoreCase);
        if (display)
        {
            foreach (var op in document
                         .Descendants(mathMlNamespace + "mo")
                         .Where(element =>
                             !string.IsNullOrEmpty(element.Value)
                             && element.Value.All(character => NaryCharacters.IndexOf(character) >= 0)
                             && (element.Parent is null || !limitNames.Contains(element.Parent.Name)))
                         .ToList())
            {
                var argument = op.ElementsAfterSelf().FirstOrDefault();
                var syntheticLimit = new XElement(
                    mathMlNamespace + "msub",
                    new XElement(op),
                    new XElement(mathMlNamespace + "mrow"));
                op.ReplaceWith(syntheticLimit);
                if (argument is null)
                {
                    syntheticLimit.AddAfterSelf(
                        new XElement(
                            mathMlNamespace + "mrow",
                            new XElement(mathMlNamespace + "mspace", new XAttribute("width", "0em"))));
                }
                else if (argument.Name != mathMlNamespace + "mrow"
                         && argument.Name != mathMlNamespace + "mstyle")
                {
                    argument.ReplaceWith(new XElement(mathMlNamespace + "mrow", argument));
                }
            }
        }

        foreach (var limit in document.Descendants().Where(element => limitNames.Contains(element.Name)).ToList())
        {
            var op = limit.Elements().FirstOrDefault();
            if (op?.Name != mathMlNamespace + "mo"
                || string.IsNullOrEmpty(op.Value)
                || op.Value.Any(character => NaryCharacters.IndexOf(character) < 0))
                continue;
            var argument = limit.ElementsAfterSelf().FirstOrDefault();
            if (argument is null
                || argument.Name == mathMlNamespace + "mrow"
                || argument.Name == mathMlNamespace + "mstyle")
                continue;

            // Office's MML2OMML.XSL recognizes an n-ary operand only when the
            // immediately following sibling is mrow or mstyle. MathJax emits a
            // valid flat Presentation MathML sequence (for example
            // munderover + mi), which Office otherwise converts to <m:e/> and
            // Word displays as a dotted placeholder box.
            argument.ReplaceWith(new XElement(mathMlNamespace + "mrow", argument));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    internal static (string MathMl, IReadOnlyList<string> NaryCharacters)
        ReplaceExtendedIntegralsWithOfficePlaceholders(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            return (mathMl, Array.Empty<string>());

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var operators = document
            .Descendants(presentationMath + "mo")
            .Where(element =>
                element.Value.Length == 1
                && NaryCharacters.IndexOf(element.Value[0]) >= 0)
            .ToArray();
        var characters = operators.Select(element => element.Value).ToArray();

        foreach (var op in operators)
        {
            if (ExtendedIntegralCharacters.IndexOf(op.Value[0]) >= 0)
            {
                // Office's MML2OMML transform knows how to attach limits and
                // the following operand to a standard integral. Use it only as
                // a structural placeholder; the exact extended character is
                // restored in OMML immediately after the transform.
                op.Value = "∫";
            }
        }

        return (document.ToString(SaveOptions.DisableFormatting), characters);
    }

    internal static string RestoreExtendedIntegralCharacters(
        string omml,
        IReadOnlyList<string> sourceNaryCharacters)
    {
        if (sourceNaryCharacters.Count == 0
            || !sourceNaryCharacters.Any(character =>
                character.Length == 1
                && ExtendedIntegralCharacters.IndexOf(character[0]) >= 0))
            return omml;

        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        var naries = document.Descendants(math + "nary").ToArray();
        if (naries.Length != sourceNaryCharacters.Count)
        {
            throw new InvalidDataException(
                "Office changed the number of n-ary operators while converting extended integrals. "
                + $"MathML={sourceNaryCharacters.Count}; OMML={naries.Length}.");
        }

        for (var index = 0; index < sourceNaryCharacters.Count; index++)
        {
            var sourceCharacter = sourceNaryCharacters[index];
            if (sourceCharacter.Length != 1
                || ExtendedIntegralCharacters.IndexOf(sourceCharacter[0]) < 0)
                continue;

            var properties = naries[index].Element(math + "naryPr");
            if (properties is null)
            {
                properties = new XElement(math + "naryPr");
                naries[index].AddFirst(properties);
            }
            var character = properties.Element(math + "chr");
            if (character is null)
            {
                character = new XElement(math + "chr");
                properties.AddFirst(character);
            }
            character.SetAttributeValue(math + "val", sourceCharacter);
        }

        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    private static bool IsBlockMathMl(string mathMl)
    {
        try
        {
            using var text = new StringReader(mathMl);
            using var reader = XmlReader.Create(text, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                MaxCharactersInDocument = 4_000_000,
            });
            var document = XDocument.Load(reader, LoadOptions.None);
            return string.Equals(
                document.Root?.Attribute("display")?.Value,
                "block",
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    internal static string NormalizeExplicitTableColumnAlignment(
        string omml,
        string mathMl)
    {
        if (string.IsNullOrWhiteSpace(omml) || string.IsNullOrWhiteSpace(mathMl))
            return omml;

        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        XNamespace officeMath = MathNamespace;
        var sourceDocument = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var targetDocument = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceTables = sourceDocument
            .Descendants(presentationMath + "mtable")
            .Where(table => !string.IsNullOrWhiteSpace(table.Attribute("columnalign")?.Value))
            .Select(table =>
            {
                var rows = table.Elements(presentationMath + "mtr").ToArray();
                var columnCount = rows.Length == 0
                    ? 0
                    : rows.Max(row => row.Elements(presentationMath + "mtd").Count());
                var raw = (table.Attribute("columnalign")?.Value ?? string.Empty)
                    .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (columnCount <= 0 || raw.Length == 0) return null;
                var alignments = Enumerable.Range(0, columnCount)
                    .Select(index => NormalizeMathMlColumnAlignment(
                        raw[Math.Min(index, raw.Length - 1)]))
                    .ToArray();
                return new
                {
                    RowCount = rows.Length,
                    ColumnCount = columnCount,
                    Alignments = alignments,
                };
            })
            .Where(table => table is not null)
            .ToArray();
        if (sourceTables.Length == 0) return omml;

        var matrices = targetDocument.Descendants(officeMath + "m").ToList();
        var used = new HashSet<XElement>();
        foreach (var sourceTable in sourceTables)
        {
            var target = matrices.FirstOrDefault(matrix =>
            {
                if (used.Contains(matrix)) return false;
                var rows = matrix.Elements(officeMath + "mr").ToArray();
                if (rows.Length != sourceTable!.RowCount) return false;
                return rows.All(row =>
                    row.Elements(officeMath + "e").Count() == sourceTable.ColumnCount);
            });
            if (target is null) continue;
            used.Add(target);

            var properties = target.Element(officeMath + "mPr");
            if (properties is null)
            {
                properties = new XElement(officeMath + "mPr");
                target.AddFirst(properties);
            }
            var columns = new XElement(officeMath + "mcs");
            foreach (var alignment in sourceTable!.Alignments)
            {
                columns.Add(
                    new XElement(
                        officeMath + "mc",
                        new XElement(
                            officeMath + "mcPr",
                            new XElement(
                                officeMath + "count",
                                new XAttribute(officeMath + "val", "1")),
                            new XElement(
                                officeMath + "mcJc",
                                new XAttribute(officeMath + "val", alignment)))));
            }
            var existing = properties.Element(officeMath + "mcs");
            if (existing is null) properties.Add(columns);
            else existing.ReplaceWith(columns);
        }

        return targetDocument.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    private static string NormalizeMathMlColumnAlignment(string value)
    {
        return value.Trim().ToLowerInvariant() switch
        {
            "left" => "left",
            "right" => "right",
            _ => "center",
        };
    }

    internal static string NormalizeDisplayNaryOmml(string omml, bool display)
    {
        if (!display) return omml;
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var nary in document.Descendants(math + "nary"))
        {
            var properties = nary.Element(math + "naryPr");
            if (properties is null)
            {
                properties = new XElement(math + "naryPr");
                nary.AddFirst(properties);
            }
            var grow = properties.Element(math + "grow");
            if (grow is null)
            {
                grow = new XElement(math + "grow");
                properties.Add(grow);
            }
            grow.SetAttributeValue(math + "val", "1");

            // The synthetic empty limit used to make a bare display integral
            // become a growing native n-ary operator must be explicitly hidden.
            // Without m:subHide/m:supHide Word renders the empty limit slot as a
            // dotted placeholder box below or above the operator.
            SetNaryLimitVisibility(
                properties,
                math + "subHide",
                !HasNaryLimitContent(nary.Element(math + "sub")));
            SetNaryLimitVisibility(
                properties,
                math + "supHide",
                !HasNaryLimitContent(nary.Element(math + "sup")));
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    private static bool HasNaryLimitContent(XElement? limit)
    {
        if (limit is null) return false;
        return limit
            .DescendantsAndSelf()
            .Where(element => element.Name.LocalName == "t")
            .Any(element => !string.IsNullOrWhiteSpace(element.Value));
    }

    private static void SetNaryLimitVisibility(
        XElement properties,
        XName propertyName,
        bool hidden)
    {
        var property = properties.Element(propertyName);
        if (!hidden)
        {
            property?.Remove();
            return;
        }
        if (property is null)
        {
            property = new XElement(propertyName);
            properties.Add(property);
        }
        property.SetAttributeValue(XName.Get("val", MathNamespace), "1");
    }

    internal static string TransformOmmlToMathMl(string wordOpenXml, bool display)
    {
        var omml = ExtractSingleOMath(wordOpenXml);
        var transform = GetOmmlToMathMlTransform();
        var inputSettings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = 4_000_000,
        };
        var outputSettings = transform.OutputSettings?.Clone() ?? new XmlWriterSettings();
        outputSettings.OmitXmlDeclaration = true;
        outputSettings.Encoding = new UTF8Encoding(false);
        using var sourceText = new StringReader(omml);
        using var source = XmlReader.Create(sourceText, inputSettings);
        using var outputText = new StringWriter();
        using (var output = XmlWriter.Create(outputText, outputSettings))
            transform.Transform(source, output);
        var transformed = outputText.ToString();
        using var transformedText = new StringReader(transformed);
        using var transformedReader = XmlReader.Create(transformedText, inputSettings);
        var document = XDocument.Load(transformedReader, LoadOptions.None);
        var root = document.Root?.Name.LocalName == "math"
            ? document.Root
            : document.Descendants().FirstOrDefault(element => element.Name.LocalName == "math");
        if (root is null)
            throw new InvalidDataException("Office OMML conversion did not produce a MathML math node.");
        root.SetAttributeValue("display", display ? "block" : "inline");
        return root.ToString(SaveOptions.DisableFormatting);
    }

    internal static string ComputeOmmlFingerprint(string wordOpenXml)
    {
        var normalized = ExtractSingleOMath(wordOpenXml);
        var document = XDocument.Parse(normalized, LoadOptions.PreserveWhitespace);
        XNamespace word = WordNamespace;

        // Word stores the visible math size in ordinary run properties. Font
        // size is presentation state, not formula content: changing 14 pt to
        // 18 pt must not force an OMML -> MathML -> LaTeX source refresh.
        document
            .Descendants()
            .Where(element => element.Name == word + "sz" || element.Name == word + "szCs")
            .Remove();

        normalized = document.Root?.ToString(SaveOptions.DisableFormatting) ?? normalized;
        using var hash = SHA256.Create();
        var bytes = hash.ComputeHash(Encoding.UTF8.GetBytes(normalized));
        return string.Concat(bytes.Select(value => value.ToString("x2")));
    }

    internal static string ExtractSingleOMath(string omml)
    {
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException("Office produced an empty OMML transformation.");
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(omml);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.None);
        XNamespace math = MathNamespace;
        var equation = document.Root?.Name == math + "oMath"
            ? document.Root
            : document.Descendants(math + "oMath").FirstOrDefault();
        if (equation is null)
            throw new InvalidDataException("Office MathML conversion did not produce an m:oMath node.");
        return equation.ToString(SaveOptions.DisableFormatting);
    }

    internal static string BuildDocumentXml(
        string omml,
        bool includeLeadingTab = false,
        bool forceInline = false)
    {
        var equation = ExtractSingleOMath(omml);
        var prefix = includeLeadingTab ? "<w:r><w:tab/></w:r>" : string.Empty;
        if (forceInline)
        {
            // Surround the source equation with ordinary runs so Word opens it
            // as an inline OMath. Only sourceMath.Range is copied later; these
            // sentinels never enter the target document.
            prefix += "<w:r><w:t>L</w:t></w:r>";
        }
        var suffix = forceInline ? "<w:r><w:t>R</w:t></w:r>" : string.Empty;
        return $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + $"<w:body><w:p>{prefix}{equation}{suffix}</w:p><w:sectPr/></w:body></w:document>";
    }

    internal static string ResolveTransformPath() => ResolveTransformPath("MML2OMML.XSL");

    internal static string ResolveReverseTransformPath() => ResolveTransformPath("OMML2MML.XSL");

    private static string ResolveTransformPath(string fileName)
    {
        var candidates = new List<string>();
        AddCandidateRoot(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), fileName);
        AddCandidateRoot(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), fileName);
        AddCandidateRoot(candidates, Environment.GetEnvironmentVariable("ProgramW6432"), fileName);
        AddCandidateRoot(candidates, AppContext.BaseDirectory, fileName);
        foreach (var candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (File.Exists(candidate)) return candidate;
        }
        throw new FileNotFoundException(
            $"Unable to locate Office {fileName}. Repair Microsoft Word or reinstall the Office integration.");
    }

    private static void AddCandidateRoot(List<string> candidates, string? root, string fileName)
    {
        if (string.IsNullOrWhiteSpace(root)) return;
        candidates.Add(Path.Combine(root, "Microsoft Office", "root", "Office16", fileName));
        candidates.Add(Path.Combine(root, "Office16", fileName));
        candidates.Add(Path.Combine(root, fileName));
    }

    private static XslCompiledTransform GetTransform()
    {
        lock (TransformLock)
        {
            if (_mathMlToOmml is not null) return _mathMlToOmml;
            _mathMlToOmml = LoadTransform(ResolveTransformPath());
            return _mathMlToOmml;
        }
    }

    private static XslCompiledTransform GetOmmlToMathMlTransform()
    {
        lock (TransformLock)
        {
            if (_ommlToMathMl is not null) return _ommlToMathMl;
            _ommlToMathMl = LoadTransform(ResolveReverseTransformPath());
            return _ommlToMathMl;
        }
    }

    private static XslCompiledTransform LoadTransform(string path)
    {
        var transform = new XslCompiledTransform(enableDebug: false);
        transform.Load(
            path,
            new XsltSettings(enableDocumentFunction: false, enableScript: false),
            null);
        return transform;
    }

    private static string CreateTemporaryDocx(
        string omml,
        bool includeLeadingTab,
        bool forceInline)
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            $"visualtex-omml-{Guid.NewGuid():N}.docx");
        using var stream = new FileStream(
            path,
            FileMode.CreateNew,
            FileAccess.ReadWrite,
            FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false);
        WriteEntry(
            archive,
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "</Types>");
        WriteEntry(
            archive,
            "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships>");
        WriteEntry(
            archive,
            "word/document.xml",
            BuildDocumentXml(omml, includeLeadingTab, forceInline));
        return path;
    }

    private static void WriteEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(content);
    }

    private static OMath? FindMathAtPosition(Document document, int position)
    {
        OMaths? maths = null;
        OMath? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    if (range.Start < position) continue;
                    var distance = range.Start - position;
                    if (distance > 16 || distance >= bestDistance) continue;
                    Release(best);
                    best = math;
                    math = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            return best;
        }
        finally { Release(maths); }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
