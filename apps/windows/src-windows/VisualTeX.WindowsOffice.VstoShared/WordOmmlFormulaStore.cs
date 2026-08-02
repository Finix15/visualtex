using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlFormulaStore
{
    internal const string NamespaceUri = "urn:visualtex:word-omml:1";
    internal const string BookmarkPrefix = "VTOMML_";
    private const string InlineBaselineBookmarkPrefix = "VTBL_";

    private static readonly XNamespace VisualTeXNamespace = NamespaceUri;
    // Keep the static initializer independent of the Office PIAs so pure XML
    // metadata tests can load this type without an installed Office assembly.
    private static readonly ConditionalWeakTable<object, DocumentMetadataCache>
        MetadataCaches = new();

    private sealed class CachedMetadataPart
    {
        internal string PartId { get; set; } = string.Empty;
        internal FormulaMetadata Metadata { get; set; } = new();
    }

    private sealed class DocumentMetadataCache
    {
        internal object Gate { get; } = new();
        internal bool Hydrated { get; set; }
        internal Dictionary<string, CachedMetadataPart> Entries { get; } =
            new(StringComparer.OrdinalIgnoreCase);
    }

    internal static string BookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException("VisualTeX OMML formulaId must be a UUID.");
        return BookmarkPrefix + parsed.ToString("N");
    }

    internal static bool TryGetFormulaId(Bookmark? bookmark, out string formulaId)
    {
        formulaId = string.Empty;
        if (bookmark is null) return false;
        string name;
        try { name = bookmark.Name ?? string.Empty; }
        catch { return false; }
        if (!name.StartsWith(BookmarkPrefix, StringComparison.Ordinal)) return false;
        var candidate = name.Substring(BookmarkPrefix.Length);
        if (!Guid.TryParseExact(candidate, "N", out var parsed)) return false;
        formulaId = parsed.ToString();
        return true;
    }

    internal static Bookmark? FindAtRange(Document document, Range selectionRange)
    {
        var direct = FindAtRangeFast(document, selectionRange);
        if (direct is not null) return direct;

        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            Bookmark? best = null;
            var bestLength = int.MaxValue;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? equationRange = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryGetFormulaId(bookmark, out _)) continue;
                    try { equationRange = GetEquationRange(bookmark); }
                    catch { continue; }
                    var containsCaret = selectionRange.Start == selectionRange.End
                        && selectionRange.Start >= equationRange.Start
                        && selectionRange.Start <= equationRange.End;
                    var overlaps = selectionRange.Start < equationRange.End
                        && selectionRange.End > equationRange.Start;
                    if (!containsCaret && !overlaps) continue;
                    var length = Math.Max(0, equationRange.End - equationRange.Start);
                    if (length >= bestLength) continue;
                    Release(best);
                    best = bookmark;
                    bookmark = null;
                    bestLength = length;
                }
                finally
                {
                    Release(equationRange);
                    Release(bookmark);
                }
            }
            return best;
        }
        finally { Release(bookmarks); }
    }

    private static Bookmark? FindAtRangeFast(
        Document document,
        Range selectionRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? equationRange = null;
        Range? content = null;
        Range? probe = null;
        Bookmarks? bookmarks = null;
        try
        {
            content = document.Content;
            object start = Math.Max(content.Start, selectionRange.Start - 2);
            object end = Math.Min(content.End, selectionRange.End + 2);
            probe = document.Range(ref start, ref end);
            bookmarks = probe.Bookmarks;

            // VisualTeX OMML bookmarks are collapsed anchors immediately before
            // the equation. Looking only inside this tiny local probe avoids the
            // previous O(bookmarks × all OMaths) scan on documents with dozens of
            // formulas.
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? bookmarkRange = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryGetFormulaId(bookmark, out _)) continue;
                    bookmarkRange = bookmark.Range;
                    if (bookmarkRange.Start >= selectionRange.Start - 2
                        && bookmarkRange.Start <= selectionRange.End + 1)
                    {
                        var result = bookmark;
                        bookmark = null;
                        return result;
                    }
                }
                catch { }
                finally
                {
                    Release(bookmarkRange);
                    Release(bookmark);
                }
            }

            // Some Word selections expose only a caret inside the OMath. Resolve
            // that one local OMath and compare its start with the nearby anchors;
            // never call GetEquationRange() for every bookmark here.
            maths = selectionRange.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            equationRange = math.Range;
            Release(bookmarks);
            bookmarks = null;
            Release(probe);
            probe = null;
            object anchorProbeStart = Math.Max(content.Start, equationRange.Start - 8);
            object anchorProbeEnd = Math.Min(content.End, equationRange.Start + 3);
            probe = document.Range(ref anchorProbeStart, ref anchorProbeEnd);
            bookmarks = probe.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? bookmarkRange = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryGetFormulaId(bookmark, out _)) continue;
                    bookmarkRange = bookmark.Range;
                    var distance = equationRange.Start - bookmarkRange.Start;
                    if (distance < 0 || distance > 2) continue;
                    var result = bookmark;
                    bookmark = null;
                    return result;
                }
                catch { }
                finally
                {
                    Release(bookmarkRange);
                    Release(bookmark);
                }
            }
            return null;
        }
        catch { return null; }
        finally
        {
            Release(bookmarks);
            Release(probe);
            Release(content);
            Release(equationRange);
            Release(math);
            Release(maths);
        }
    }

    internal static Bookmark? FindByFormulaId(Document document, string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = BookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return null;
            return bookmarks[name];
        }
        finally { Release(bookmarks); }
    }

    internal static IReadOnlyList<string> FormulaIds(Document document)
    {
        var result = new List<string>();
        var staleFormulaIds = new List<string>();
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? bookmarkRange = null;
                Range? equationRange = null;
                OMaths? maths = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryGetFormulaId(bookmark, out var formulaId)) continue;
                    bookmarkRange = bookmark.Range;
                    try
                    {
                        equationRange = GetEquationRange(bookmark);
                        maths = equationRange.OMaths;
                        var anchorDistance = equationRange.Start - bookmarkRange.Start;
                        if (maths.Count == 1 && anchorDistance >= 0 && anchorDistance <= 1)
                            result.Add(formulaId);
                        else
                            staleFormulaIds.Add(formulaId);
                    }
                    catch
                    {
                        // Deleting an OMML formula with Word's Delete key leaves
                        // its collapsed VisualTeX bookmark and custom XML part.
                        // Treat that anchor as stale so Update Equation Numbers
                        // can remove the old visible number and renumber the rest.
                        staleFormulaIds.Add(formulaId);
                    }
                }
                finally
                {
                    Release(maths);
                    Release(equationRange);
                    Release(bookmarkRange);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }

        foreach (var formulaId in staleFormulaIds.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            Bookmark? staleBookmark = null;
            try
            {
                staleBookmark = FindByFormulaId(document, formulaId);
                staleBookmark?.Delete();
                Delete(document, formulaId);
            }
            catch { }
            finally { Release(staleBookmark); }
        }
        return result;
    }

    internal static FormulaMetadata? TryRead(Document document, Bookmark bookmark)
    {
        if (!TryGetFormulaId(bookmark, out var formulaId)) return null;
        return TryRead(document, formulaId);
    }

    internal static FormulaMetadata? TryRead(Document document, string formulaId)
    {
        var cache = MetadataCaches.GetValue(document, _ => new DocumentMetadataCache());
        lock (cache.Gate)
        {
            if (cache.Entries.TryGetValue(formulaId, out var cached))
                return CloneMetadata(cached.Metadata);
        }

        object? part = null;
        try
        {
            part = FindPart(document, formulaId);
            if (part is null) return null;
            var partXml = (string?)((dynamic)part).XML;
            if (!TryDecodePartXml(partXml, out var metadata)
                || !string.Equals(metadata.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                return null;
            RememberPart(document, part, metadata);
            return CloneMetadata(metadata);
        }
        catch
        {
            return null;
        }
        finally { Release(part); }
    }

    internal static void Save(Document document, FormulaMetadata metadata)
    {
        metadata.Validate();
        var xml = BuildPartXml(metadata);
        object? existing = null;
        object? parts = null;
        object? added = null;
        try
        {
            existing = FindPart(document, metadata.FormulaId);
            if (existing is not null)
            {
                try
                {
                    if ((bool)((dynamic)existing).LoadXML(xml))
                    {
                        RememberPart(document, existing, metadata);
                        return;
                    }
                }
                catch
                {
                    // Some Word builds reject in-place replacement after an
                    // OMath rebuild. Fall back to add-then-delete below.
                }
            }
            parts = ((dynamic)document).CustomXMLParts;
            added = ((dynamic)parts).Add(xml);
            if (added is null)
                throw new InvalidOperationException("Word did not create the VisualTeX OMML metadata part.");

            // Word may reject CustomXMLPart.LoadXML after an OMath rebuild.
            // Add the replacement first, then remove the old part so a failed
            // update never destroys the last valid metadata copy.
            if (existing is not null) ((dynamic)existing).Delete();
            RememberPart(document, added, metadata);
        }
        finally
        {
            Release(added);
            Release(parts);
            Release(existing);
        }
    }

    internal static void Delete(Document document, string formulaId)
    {
        object? part = null;
        try
        {
            part = FindPart(document, formulaId);
            if (part is not null) ((dynamic)part).Delete();
            ForgetPart(document, formulaId);
        }
        finally { Release(part); }
    }

    internal static Bookmark Wrap(
        Document document,
        Range equationRange,
        FormulaMetadata metadata)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? anchorRange = null;
        Range? preceding = null;
        try
        {
            var anchorPosition = equationRange.Start;
            if (anchorPosition > 0)
            {
                object precedingStart = anchorPosition - 1;
                object precedingEnd = anchorPosition;
                preceding = document.Range(ref precedingStart, ref precedingEnd);
                if (string.Equals(preceding.Text, "\v", StringComparison.Ordinal))
                    anchorPosition--;
            }

            object anchorStart = anchorPosition;
            object anchorEnd = anchorPosition;
            anchorRange = document.Range(ref anchorStart, ref anchorEnd);
            bookmarks = document.Bookmarks;
            var name = BookmarkName(metadata.FormulaId);
            if (bookmarks.Exists(name)) bookmarks[name].Delete();
            bookmark = bookmarks.Add(name, anchorRange);
            var result = bookmark;
            bookmark = null;
            return result;
        }
        finally
        {
            Release(preceding);
            Release(anchorRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static Range GetEquationRange(Bookmark bookmark)
    {
        Range? bookmarkRange = null;
        Document? document = null;
        Range? content = null;
        OMaths? maths = null;
        Range? bestRange = null;
        try
        {
            bookmarkRange = bookmark.Range;
            document = bookmarkRange.Document;
            content = document.Content;
            var anchor = bookmarkRange.Start;

            // The bookmark is a collapsed anchor immediately before the OMath.
            // Word clips OMath.Range to the range used to obtain OMaths, so a
            // fixed short probe can silently truncate the tail of a long formula.
            // Expand locally until the returned equation ends before the probe;
            // this remains independent of the total formula count in the document.
            bestRange = FindAdjacentEquationRangeNearAnchor(
                document,
                content,
                anchor);

            // Preserve compatibility with anchors moved by external Word edits.
            // This slow path should be exceptional, not the normal lookup path.
            if (bestRange is null)
            {
                Release(maths);
                maths = document.OMaths;
                bestRange = FindAdjacentEquationRange(maths, anchor);
            }

            if (bestRange is null)
                throw new InvalidDataException(
                    "The VisualTeX OMML anchor is no longer adjacent to a Word equation.");
            ClampToInlineBaselineBookmark(document, bookmark, bestRange);
            var result = bestRange;
            bestRange = null;
            return result;
        }
        finally
        {
            Release(bestRange);
            Release(maths);
            Release(content);
            Release(document);
            Release(bookmarkRange);
        }
    }

    private static void ClampToInlineBaselineBookmark(
        Document document,
        Bookmark formulaBookmark,
        Range equationRange)
    {
        if (!TryGetFormulaId(formulaBookmark, out var formulaId)
            || !Guid.TryParse(formulaId, out var parsed))
            return;
        Bookmarks? bookmarks = null;
        Bookmark? baselineBookmark = null;
        Range? baselineRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = InlineBaselineBookmarkPrefix + parsed.ToString("N");
            if (!bookmarks.Exists(name)) return;
            baselineBookmark = bookmarks[name];
            baselineRange = baselineBookmark.Range;
            if (baselineRange.Start < equationRange.Start
                || baselineRange.Start > equationRange.End + 8)
                return;
            equationRange.End = Math.Min(equationRange.End, baselineRange.Start);
        }
        catch
        {
            // A stale typing bookmark must not block editing the formula itself.
        }
        finally
        {
            Release(baselineRange);
            Release(baselineBookmark);
            Release(bookmarks);
        }
    }

    private static Range? FindAdjacentEquationRangeNearAnchor(
        Document document,
        Range content,
        int anchor)
    {
        var span = 64;
        while (span <= 65_536)
        {
            Range? probe = null;
            OMaths? maths = null;
            Range? candidate = null;
            try
            {
                object probeStart = Math.Max(content.Start, anchor);
                object probeEnd = Math.Min(content.End, anchor + span);
                if ((int)probeEnd <= (int)probeStart && content.End > (int)probeStart)
                    probeEnd = (int)probeStart + 1;
                probe = document.Range(ref probeStart, ref probeEnd);
                maths = probe.OMaths;
                candidate = FindAdjacentEquationRange(maths, anchor);
                if (candidate is null) return null;

                // A range touching the probe boundary may have been clipped by
                // Word. Grow the probe and reacquire the same OMath until there
                // is at least one character of room after the equation.
                if (candidate.End < probe.End || probe.End >= content.End)
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }
            finally
            {
                Release(candidate);
                Release(maths);
                Release(probe);
            }
            span *= 4;
        }
        return null;
    }

    private static Range? FindAdjacentEquationRange(OMaths maths, int anchor)
    {
        Range? bestRange = null;
        var bestDistance = int.MaxValue;
        for (var index = 1; index <= maths.Count; index++)
        {
            OMath? math = null;
            Range? range = null;
            try
            {
                math = maths[index];
                range = math.Range;
                var containsAnchor = anchor >= range.Start && anchor <= range.End;
                var distance = containsAnchor ? 0 : range.Start - anchor;
                if (distance < 0 || distance > 8 || distance >= bestDistance) continue;
                Release(bestRange);
                bestRange = TrimToNativeMath(range);
                bestDistance = distance;
            }
            finally
            {
                Release(range);
                Release(math);
            }
        }
        return bestRange;
    }

    private static Range TrimToNativeMath(Range source)
    {
        Range? result = null;
        Range? probe = null;
        try
        {
            result = source.Duplicate;
            while (result.Start < result.End)
            {
                probe = result.Duplicate;
                probe.SetRange(result.Start, Math.Min(result.Start + 1, result.End));
                if (RangeContainsNativeMath(probe)) break;
                result.Start++;
                Release(probe);
                probe = null;
            }
            while (result.End > result.Start)
            {
                probe = result.Duplicate;
                probe.SetRange(Math.Max(result.Start, result.End - 1), result.End);
                if (RangeContainsNativeMath(probe)) break;
                result.End--;
                Release(probe);
                probe = null;
            }
            if (result.Start >= result.End)
                throw new InvalidDataException(
                    "Word returned an OMML range without native math content.");
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(probe);
            Release(result);
        }
    }

    private static bool RangeContainsNativeMath(Range range)
    {
        OMaths? maths = null;
        try
        {
            maths = range.OMaths;
            return maths.Count > 0;
        }
        catch { return false; }
        finally { Release(maths); }
    }

    internal static float EstimateHeightPoints(Bookmark bookmark)
    {
        Range? equationRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            equationRange = GetEquationRange(bookmark);
            font = equationRange.Font;
            var size = 11f;
            try { size = font.Size; } catch { }
            if (float.IsNaN(size) || float.IsInfinity(size) || size <= 0 || size > 256)
                size = 11f;
            return Math.Max(11f, size * 1.5f);
        }
        finally
        {
            Release(font);
            Release(equationRange);
        }
    }

    internal static string BuildPartXml(FormulaMetadata metadata)
    {
        metadata.Validate();
        var encoded = FormulaMetadataCodec.Encode(metadata);
        return new XDocument(
            new XElement(
                VisualTeXNamespace + "formula",
                new XAttribute("formulaId", metadata.FormulaId),
                new XElement(VisualTeXNamespace + "metadata", encoded)))
            .ToString(SaveOptions.DisableFormatting);
    }

    internal static bool TryDecodePartXml(string? xml, out FormulaMetadata metadata)
    {
        metadata = null!;
        if (string.IsNullOrWhiteSpace(xml)) return false;
        try
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
            };
            using var stringReader = new StringReader(xml);
            using var reader = XmlReader.Create(stringReader, settings);
            var document = XDocument.Load(reader, LoadOptions.None);
            var root = document.Root;
            if (root?.Name != VisualTeXNamespace + "formula") return false;
            var formulaId = (string?)root.Attribute("formulaId");
            var encoded = root.Element(VisualTeXNamespace + "metadata")?.Value;
            var decoded = FormulaMetadataCodec.Decode(encoded);
            if (decoded is null
                || !string.Equals(decoded.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                return false;
            decoded.Validate();
            metadata = decoded;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static FormulaMetadata CloneMetadata(FormulaMetadata metadata) =>
        FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(metadata))
        ?? throw new InvalidDataException("Unable to clone VisualTeX OMML metadata.");

    private static void RememberPart(
        Document document,
        object? part,
        FormulaMetadata metadata)
    {
        if (part is null) return;
        string? partId = null;
        try { partId = (string?)((dynamic)part).Id; } catch { }
        if (string.IsNullOrWhiteSpace(partId)) return;
        var cache = MetadataCaches.GetValue(document, _ => new DocumentMetadataCache());
        lock (cache.Gate)
        {
            cache.Entries[metadata.FormulaId] = new CachedMetadataPart
            {
                PartId = partId!,
                Metadata = CloneMetadata(metadata),
            };
        }
    }

    private static void ForgetPart(Document document, string formulaId)
    {
        if (!MetadataCaches.TryGetValue(document, out var cache)) return;
        lock (cache.Gate) cache.Entries.Remove(formulaId);
    }

    private static object? FindPart(Document document, string formulaId)
    {
        var cache = MetadataCaches.GetValue(document, _ => new DocumentMetadataCache());
        string? cachedPartId = null;
        lock (cache.Gate)
        {
            if (cache.Entries.TryGetValue(formulaId, out var cached))
                cachedPartId = cached.PartId;
        }

        object? parts = null;
        object? selected = null;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            if (!string.IsNullOrWhiteSpace(cachedPartId))
            {
                object? cachedPart = null;
                try
                {
                    cachedPart = ((dynamic)parts).SelectByID(cachedPartId!);
                    if (cachedPart is not null)
                    {
                        var partXml = (string?)((dynamic)cachedPart).XML;
                        if (TryDecodePartXml(partXml, out var cachedMetadata)
                            && string.Equals(
                                cachedMetadata.FormulaId,
                                formulaId,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            var result = cachedPart;
                            cachedPart = null;
                            return result;
                        }
                    }
                }
                catch { }
                finally { Release(cachedPart); }
                ForgetPart(document, formulaId);
            }

            selected = ((dynamic)parts).SelectByNamespace(NamespaceUri);
            var count = (int)((dynamic)selected).Count;
            object? matched = null;
            for (var index = 1; index <= count; index++)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)selected)[index];
                    var partXml = (string?)((dynamic)part).XML;
                    if (!TryDecodePartXml(partXml, out var metadata)) continue;
                    RememberPart(document, part, metadata);
                    if (matched is null
                        && string.Equals(
                            metadata.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        matched = part;
                        part = null;
                    }
                }
                finally { Release(part); }
            }
            lock (cache.Gate) cache.Hydrated = true;
            return matched;
        }
        finally
        {
            Release(selected);
            Release(parts);
        }
    }

    internal static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
