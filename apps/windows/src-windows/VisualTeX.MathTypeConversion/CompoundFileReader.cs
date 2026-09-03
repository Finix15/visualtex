using System.Text;

namespace VisualTeX.MathTypeConversion;

internal sealed class CompoundFileReader
{
    private const uint EndOfChain = 0xfffffffe;
    private const uint FreeSector = 0xffffffff;
    private const int MaximumSectors = 131072;
    private readonly byte[] _data;
    private readonly int _sectorSize;
    private readonly int _miniSectorSize;
    private readonly uint[] _fat;
    private readonly uint[] _miniFat;
    private readonly byte[] _miniStream;
    private readonly List<DirectoryEntry> _entries;

    public CompoundFileReader(byte[] data)
    {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        if (data.Length < 512 || data.Length > MathTypeLimits.MaximumPayloadBytes)
            throw new InvalidDataException("The OLE compound file size is invalid.");
        var signature = new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
        if (!data.Take(8).SequenceEqual(signature))
            throw new InvalidDataException("The embedded object is not an OLE compound file.");
        if (ReadUInt16(28) != 0xfffe)
            throw new InvalidDataException("The OLE compound file byte order is unsupported.");
        _sectorSize = CheckedSectorSize(ReadUInt16(30), 9, 12, "sector");
        _miniSectorSize = CheckedSectorSize(ReadUInt16(32), 6, 6, "mini-sector");
        var fatSectorIds = ReadFatSectorIds();
        _fat = ReadSectorTable(fatSectorIds);
        var directoryBytes = ReadChain(ReadUInt32(48), _fat, -1, false);
        _entries = ReadDirectory(directoryBytes);
        var root = _entries.FirstOrDefault(entry => entry.Type == 5)
            ?? throw new InvalidDataException("The OLE compound file has no root storage.");
        _miniStream = ReadChain(root.StartSector, _fat, CheckedLength(root.Size), false);
        var miniFatStart = ReadUInt32(60);
        var miniFatSectors = ReadUInt32(64);
        _miniFat = miniFatSectors == 0 || miniFatStart == EndOfChain
            ? Array.Empty<uint>()
            : ToUInt32Array(ReadChain(miniFatStart, _fat, checked((int)miniFatSectors * _sectorSize), false));
    }

    public byte[] ReadStream(string name)
    {
        var entry = _entries.FirstOrDefault(candidate =>
            candidate.Type == 2 && string.Equals(candidate.Name, name, StringComparison.Ordinal));
        if (entry is null) throw new InvalidDataException($"The OLE stream '{name}' is missing.");
        var length = CheckedLength(entry.Size);
        if (length == 0) return Array.Empty<byte>();
        return entry.Size < ReadUInt32(56)
            ? ReadMiniChain(entry.StartSector, length)
            : ReadChain(entry.StartSector, _fat, length, false);
    }

    public IEnumerable<byte[]> ReadAllStreams()
    {
        foreach (var entry in _entries.Where(entry => entry.Type == 2 && entry.Size > 0))
        {
            byte[] value;
            try { value = ReadStream(entry.Name); }
            catch (InvalidDataException) { continue; }
            yield return value;
        }
    }

    private List<uint> ReadFatSectorIds()
    {
        var required = ReadUInt32(44);
        if (required > MaximumSectors) throw new InvalidDataException("The OLE FAT is too large.");
        var ids = new List<uint>((int)required);
        for (var index = 0; index < 109 && ids.Count < required; index++)
        {
            var id = ReadUInt32(76 + index * 4);
            if (id != FreeSector) ids.Add(id);
        }
        var difat = ReadUInt32(68);
        var difatCount = ReadUInt32(72);
        for (var sectorIndex = 0; sectorIndex < difatCount && ids.Count < required; sectorIndex++)
        {
            var sector = ReadSector(difat);
            var entries = _sectorSize / 4 - 1;
            for (var index = 0; index < entries && ids.Count < required; index++)
            {
                var id = BitConverter.ToUInt32(sector, index * 4);
                if (id != FreeSector) ids.Add(id);
            }
            difat = BitConverter.ToUInt32(sector, _sectorSize - 4);
            if (difat == EndOfChain) break;
        }
        if (ids.Count != required) throw new InvalidDataException("The OLE FAT sector list is incomplete.");
        return ids;
    }

    private uint[] ReadSectorTable(IEnumerable<uint> sectorIds)
    {
        using var output = new MemoryStream();
        foreach (var id in sectorIds)
        {
            var sector = ReadSector(id);
            output.Write(sector, 0, sector.Length);
        }
        return ToUInt32Array(output.ToArray());
    }

    private byte[] ReadChain(uint start, uint[] table, int requestedLength, bool mini)
    {
        if (start == EndOfChain && requestedLength == 0) return Array.Empty<byte>();
        using var output = new MemoryStream();
        var current = start;
        var seen = new HashSet<uint>();
        while (current != EndOfChain && (requestedLength < 0 || output.Length < requestedLength))
        {
            if (current >= table.Length || !seen.Add(current) || seen.Count > MaximumSectors)
                throw new InvalidDataException("The OLE sector chain is invalid.");
            byte[] sector;
            if (mini)
            {
                var offset = checked((long)current * _miniSectorSize);
                if (offset < 0 || offset + _miniSectorSize > _miniStream.Length)
                    throw new InvalidDataException("The OLE mini-sector is outside the mini stream.");
                sector = new byte[_miniSectorSize];
                Buffer.BlockCopy(_miniStream, (int)offset, sector, 0, _miniSectorSize);
            }
            else sector = ReadSector(current);
            var count = requestedLength < 0
                ? sector.Length
                : Math.Min(sector.Length, requestedLength - checked((int)output.Length));
            output.Write(sector, 0, count);
            current = table[current];
        }
        if (requestedLength >= 0 && output.Length < requestedLength)
            throw new InvalidDataException("The OLE sector chain ended before the declared stream length.");
        return output.ToArray();
    }

    private byte[] ReadMiniChain(uint start, int length)
    {
        if (_miniFat.Length == 0) throw new InvalidDataException("The OLE mini FAT is missing.");
        return ReadChain(start, _miniFat, length, true);
    }

    private byte[] ReadSector(uint id)
    {
        var offset = checked(((long)id + 1) * _sectorSize);
        if (id >= MaximumSectors || offset < 0 || offset + _sectorSize > _data.Length)
            throw new InvalidDataException("The OLE sector is outside the file.");
        var result = new byte[_sectorSize];
        Buffer.BlockCopy(_data, (int)offset, result, 0, _sectorSize);
        return result;
    }

    private List<DirectoryEntry> ReadDirectory(byte[] bytes)
    {
        var result = new List<DirectoryEntry>();
        for (var offset = 0; offset + 128 <= bytes.Length; offset += 128)
        {
            var nameLength = BitConverter.ToUInt16(bytes, offset + 64);
            var type = bytes[offset + 66];
            if (type == 0 || nameLength < 2 || nameLength > 64 || (nameLength & 1) != 0) continue;
            var name = Encoding.Unicode.GetString(bytes, offset, nameLength - 2);
            var start = BitConverter.ToUInt32(bytes, offset + 116);
            var size = BitConverter.ToUInt64(bytes, offset + 120);
            result.Add(new DirectoryEntry(name, type, start, size));
        }
        return result;
    }

    private static uint[] ToUInt32Array(byte[] bytes)
    {
        var values = new uint[bytes.Length / 4];
        for (var index = 0; index < values.Length; index++)
            values[index] = BitConverter.ToUInt32(bytes, index * 4);
        return values;
    }

    private static int CheckedSectorSize(ushort shift, int minimum, int maximum, string label)
    {
        if (shift < minimum || shift > maximum)
            throw new InvalidDataException($"The OLE {label} size is unsupported.");
        return 1 << shift;
    }

    private static int CheckedLength(ulong value)
    {
        if (value > MathTypeLimits.MaximumPayloadBytes)
            throw new InvalidDataException("The OLE stream is too large.");
        return checked((int)value);
    }

    private ushort ReadUInt16(int offset) => BitConverter.ToUInt16(_data, offset);
    private uint ReadUInt32(int offset) => BitConverter.ToUInt32(_data, offset);

    private sealed class DirectoryEntry
    {
        public DirectoryEntry(string name, byte type, uint startSector, ulong size)
        {
            Name = name;
            Type = type;
            StartSector = startSector;
            Size = size;
        }

        public string Name { get; }
        public byte Type { get; }
        public uint StartSector { get; }
        public ulong Size { get; }
    }
}

internal static class MathTypeLimits
{
    public const int MaximumPayloadBytes = 16 * 1024 * 1024;
    public const int MaximumDepth = 256;
    public const int MaximumRecords = 1_000_000;
}
