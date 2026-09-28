using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PptxViewer
{
    internal sealed class CompoundFileDirectoryEntry
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public byte ObjectType { get; set; }
        public int LeftSiblingId { get; set; }
        public int RightSiblingId { get; set; }
        public int ChildId { get; set; }
        public uint StartSector { get; set; }
        public long StreamSize { get; set; }

        public bool IsStorage
        {
            get { return ObjectType == 1 || ObjectType == 5; }
        }

        public bool IsStream
        {
            get { return ObjectType == 2; }
        }

        public CompoundFileDirectoryEntry()
        {
            Name = string.Empty;
            LeftSiblingId = -1;
            RightSiblingId = -1;
            ChildId = -1;
        }
    }

    internal sealed class CompoundFileReader : IDisposable
    {
        private const uint FreeSector = 0xFFFFFFFF;
        private const uint EndOfChain = 0xFFFFFFFE;
        private const uint FatSector = 0xFFFFFFFD;
        private const uint DifatSector = 0xFFFFFFFC;
        private const uint NoStream = 0xFFFFFFFF;

        private readonly FileStream stream;
        private readonly List<uint> fat = new List<uint>();
        private readonly List<uint> miniFat = new List<uint>();
        private readonly List<CompoundFileDirectoryEntry> directory =
            new List<CompoundFileDirectoryEntry>();

        private int sectorSize;
        private int miniSectorSize;
        private uint miniStreamCutoff;
        private ushort majorVersion;
        private byte[] rootMiniStream;
        private CompoundFileDirectoryEntry rootEntry;
        private bool disposed;

        public CompoundFileReader(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("Compound File was not found.", path);

            stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite);

            try
            {
                Parse();
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        public IList<string> GetStreamPaths()
        {
            List<string> result = new List<string>();
            if (rootEntry == null)
                return result;

            HashSet<int> visited = new HashSet<int>();
            EnumerateChildren(rootEntry, string.Empty, result, visited);
            result.Sort(StringComparer.OrdinalIgnoreCase);
            return result;
        }

        public bool StreamExists(string path)
        {
            CompoundFileDirectoryEntry entry = FindEntry(path);
            return entry != null && entry.IsStream;
        }

        public byte[] ReadStream(string path)
        {
            CompoundFileDirectoryEntry entry = FindEntry(path);
            if (entry == null || !entry.IsStream)
                throw new FileNotFoundException("Compound File stream was not found: " + path);

            return ReadEntryStream(entry);
        }

        private void Parse()
        {
            byte[] header = ReadExactAt(0, 512);
            byte[] signature = new byte[]
            {
                0xD0, 0xCF, 0x11, 0xE0,
                0xA1, 0xB1, 0x1A, 0xE1
            };

            for (int i = 0; i < signature.Length; i++)
            {
                if (header[i] != signature[i])
                    throw new InvalidDataException("The file is not a Compound File Binary container.");
            }

            majorVersion = ReadUInt16(header, 26);
            ushort byteOrder = ReadUInt16(header, 28);
            ushort sectorShift = ReadUInt16(header, 30);
            ushort miniSectorShift = ReadUInt16(header, 32);

            if (byteOrder != 0xFFFE)
                throw new InvalidDataException("Unsupported Compound File byte order.");

            if (majorVersion != 3 && majorVersion != 4)
                throw new InvalidDataException("Unsupported Compound File major version: " + majorVersion.ToString());

            sectorSize = 1 << sectorShift;
            miniSectorSize = 1 << miniSectorShift;

            if ((sectorSize != 512 && sectorSize != 4096) || miniSectorSize != 64)
                throw new InvalidDataException("Unsupported Compound File sector size.");

            uint fatSectorCount = ReadUInt32(header, 44);
            uint firstDirectorySector = ReadUInt32(header, 48);
            miniStreamCutoff = ReadUInt32(header, 56);
            uint firstMiniFatSector = ReadUInt32(header, 60);
            uint miniFatSectorCount = ReadUInt32(header, 64);
            uint firstDifatSector = ReadUInt32(header, 68);
            uint difatSectorCount = ReadUInt32(header, 72);

            List<uint> fatSectorIds = ReadDifat(
                header,
                fatSectorCount,
                firstDifatSector,
                difatSectorCount);

            LoadFat(fatSectorIds);
            LoadDirectory(firstDirectorySector);
            LoadMiniFat(firstMiniFatSector, miniFatSectorCount);
            LoadRootMiniStream();
        }

        private List<uint> ReadDifat(
            byte[] header,
            uint expectedFatSectors,
            uint firstDifatSector,
            uint difatSectorCount)
        {
            List<uint> result = new List<uint>();

            for (int i = 0; i < 109; i++)
            {
                uint sector = ReadUInt32(header, 76 + i * 4);
                if (IsRegularSector(sector))
                    result.Add(sector);
            }

            uint current = firstDifatSector;
            uint traversed = 0;
            int entriesPerSector = sectorSize / 4 - 1;
            HashSet<uint> visited = new HashSet<uint>();

            while (IsRegularSector(current) && traversed < difatSectorCount)
            {
                if (!visited.Add(current))
                    throw new InvalidDataException("Loop detected in DIFAT chain.");

                byte[] data = ReadSector(current);

                for (int i = 0; i < entriesPerSector; i++)
                {
                    uint fatSectorId = ReadUInt32(data, i * 4);
                    if (IsRegularSector(fatSectorId))
                        result.Add(fatSectorId);
                }

                current = ReadUInt32(data, entriesPerSector * 4);
                traversed++;
            }

            if (expectedFatSectors > 0 && result.Count < expectedFatSectors)
                throw new InvalidDataException("Compound File FAT sector list is truncated.");

            if (expectedFatSectors > 0 && result.Count > expectedFatSectors)
                result.RemoveRange((int)expectedFatSectors, result.Count - (int)expectedFatSectors);

            return result;
        }

        private void LoadFat(List<uint> fatSectorIds)
        {
            int entriesPerSector = sectorSize / 4;

            for (int i = 0; i < fatSectorIds.Count; i++)
            {
                byte[] data = ReadSector(fatSectorIds[i]);
                for (int j = 0; j < entriesPerSector; j++)
                    fat.Add(ReadUInt32(data, j * 4));
            }

            if (fat.Count == 0)
                throw new InvalidDataException("Compound File contains no FAT entries.");
        }

        private void LoadDirectory(uint firstDirectorySector)
        {
            byte[] data = ReadRegularChain(firstDirectorySector, -1, 0);

            for (int offset = 0; offset + 128 <= data.Length; offset += 128)
            {
                CompoundFileDirectoryEntry entry = new CompoundFileDirectoryEntry();
                entry.Id = directory.Count;

                ushort nameLength = ReadUInt16(data, offset + 64);
                if (nameLength >= 2 && nameLength <= 64)
                {
                    int charBytes = nameLength - 2;
                    entry.Name = Encoding.Unicode.GetString(data, offset, charBytes);
                }

                entry.ObjectType = data[offset + 66];
                entry.LeftSiblingId = NormalizeDirectoryId(ReadUInt32(data, offset + 68));
                entry.RightSiblingId = NormalizeDirectoryId(ReadUInt32(data, offset + 72));
                entry.ChildId = NormalizeDirectoryId(ReadUInt32(data, offset + 76));
                entry.StartSector = ReadUInt32(data, offset + 116);

                ulong streamSize = ReadUInt64(data, offset + 120);
                if (majorVersion == 3)
                    streamSize &= 0xFFFFFFFFUL;

                if (streamSize > int.MaxValue)
                    throw new InvalidDataException("Compound File stream is too large for this reader.");

                entry.StreamSize = (long)streamSize;
                directory.Add(entry);

                if (entry.ObjectType == 5 && rootEntry == null)
                    rootEntry = entry;
            }

            if (rootEntry == null)
                throw new InvalidDataException("Compound File root storage is missing.");
        }

        private void LoadMiniFat(uint firstMiniFatSector, uint miniFatSectorCount)
        {
            if (miniFatSectorCount == 0 || !IsRegularSector(firstMiniFatSector))
                return;

            long expectedBytes = (long)miniFatSectorCount * sectorSize;
            byte[] data = ReadRegularChain(firstMiniFatSector, expectedBytes, (int)miniFatSectorCount + 8);

            for (int offset = 0; offset + 4 <= data.Length; offset += 4)
                miniFat.Add(ReadUInt32(data, offset));
        }

        private void LoadRootMiniStream()
        {
            if (rootEntry == null || rootEntry.StreamSize <= 0 || !IsRegularSector(rootEntry.StartSector))
            {
                rootMiniStream = new byte[0];
                return;
            }

            rootMiniStream = ReadRegularChain(
                rootEntry.StartSector,
                rootEntry.StreamSize,
                0);
        }

        private CompoundFileDirectoryEntry FindEntry(string path)
        {
            if (rootEntry == null || string.IsNullOrEmpty(path))
                return null;

            string normalized = path.Replace('\\', '/').Trim('/');
            if (normalized.Length == 0)
                return rootEntry;

            string[] pieces = normalized.Split('/');
            CompoundFileDirectoryEntry current = rootEntry;

            for (int i = 0; i < pieces.Length; i++)
            {
                current = FindChild(current, pieces[i]);
                if (current == null)
                    return null;
            }

            return current;
        }

        private CompoundFileDirectoryEntry FindChild(
            CompoundFileDirectoryEntry storage,
            string name)
        {
            if (storage == null || storage.ChildId < 0)
                return null;

            List<int> ids = new List<int>();
            HashSet<int> visited = new HashSet<int>();
            CollectSiblingTree(storage.ChildId, ids, visited);

            for (int i = 0; i < ids.Count; i++)
            {
                CompoundFileDirectoryEntry child = GetDirectoryEntry(ids[i]);
                if (child != null && string.Equals(
                        child.Name,
                        name,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return child;
                }
            }

            return null;
        }

        private void EnumerateChildren(
            CompoundFileDirectoryEntry storage,
            string prefix,
            List<string> result,
            HashSet<int> visited)
        {
            if (storage == null || storage.ChildId < 0)
                return;

            List<int> ids = new List<int>();
            CollectSiblingTree(storage.ChildId, ids, visited);

            for (int i = 0; i < ids.Count; i++)
            {
                CompoundFileDirectoryEntry child = GetDirectoryEntry(ids[i]);
                if (child == null || string.IsNullOrEmpty(child.Name))
                    continue;

                string path = string.IsNullOrEmpty(prefix)
                    ? child.Name
                    : prefix + "/" + child.Name;

                if (child.IsStream)
                    result.Add(path);
                else if (child.IsStorage)
                    EnumerateChildren(child, path, result, visited);
            }
        }

        private void CollectSiblingTree(
            int id,
            List<int> output,
            HashSet<int> visited)
        {
            if (id < 0 || id >= directory.Count || !visited.Add(id))
                return;

            CompoundFileDirectoryEntry entry = directory[id];
            CollectSiblingTree(entry.LeftSiblingId, output, visited);
            output.Add(id);
            CollectSiblingTree(entry.RightSiblingId, output, visited);
        }

        private CompoundFileDirectoryEntry GetDirectoryEntry(int id)
        {
            return id >= 0 && id < directory.Count ? directory[id] : null;
        }

        private byte[] ReadEntryStream(CompoundFileDirectoryEntry entry)
        {
            if (entry == null || entry.StreamSize <= 0)
                return new byte[0];

            if (entry.ObjectType != 5 &&
                entry.StreamSize < miniStreamCutoff &&
                miniFat.Count > 0 &&
                rootMiniStream != null &&
                rootMiniStream.Length > 0)
            {
                return ReadMiniChain(entry.StartSector, entry.StreamSize);
            }

            return ReadRegularChain(entry.StartSector, entry.StreamSize, 0);
        }

        private byte[] ReadMiniChain(uint startSector, long size)
        {
            if (size <= 0)
                return new byte[0];
            if (size > int.MaxValue)
                throw new InvalidDataException("Mini stream is too large.");

            byte[] result = new byte[(int)size];
            int written = 0;
            uint current = startSector;
            HashSet<uint> visited = new HashSet<uint>();

            while (written < result.Length && IsRegularSector(current))
            {
                if (!visited.Add(current))
                    throw new InvalidDataException("Loop detected in MiniFAT chain.");
                if (current >= miniFat.Count)
                    throw new InvalidDataException("MiniFAT sector index is out of range.");

                long offset = (long)current * miniSectorSize;
                if (offset < 0 || offset >= rootMiniStream.Length)
                    throw new InvalidDataException("Mini stream sector is out of range.");

                int available = Math.Min(
                    miniSectorSize,
                    rootMiniStream.Length - (int)offset);
                int copy = Math.Min(available, result.Length - written);
                Buffer.BlockCopy(rootMiniStream, (int)offset, result, written, copy);
                written += copy;

                current = miniFat[(int)current];
            }

            if (written < result.Length)
                throw new InvalidDataException("Mini stream ended before the declared size.");

            return result;
        }

        private byte[] ReadRegularChain(uint startSector, long declaredSize, int maxSectors)
        {
            if (!IsRegularSector(startSector))
                return new byte[0];

            List<byte[]> sectors = new List<byte[]>();
            uint current = startSector;
            HashSet<uint> visited = new HashSet<uint>();
            long total = 0;
            int hardLimit = maxSectors > 0
                ? maxSectors
                : Math.Max(16, fat.Count + 8);

            while (IsRegularSector(current))
            {
                if (sectors.Count >= hardLimit)
                    throw new InvalidDataException("Compound File stream chain exceeds the safety limit.");
                if (!visited.Add(current))
                    throw new InvalidDataException("Loop detected in FAT chain.");
                if (current >= fat.Count)
                    throw new InvalidDataException("FAT sector index is out of range.");

                byte[] sector = ReadSector(current);
                sectors.Add(sector);
                total += sector.Length;

                if (declaredSize >= 0 && total >= declaredSize)
                    break;

                current = fat[(int)current];
            }

            long outputLength = declaredSize >= 0
                ? declaredSize
                : total;

            if (outputLength < 0 || outputLength > int.MaxValue)
                throw new InvalidDataException("Compound File stream is too large.");
            if (declaredSize >= 0 && total < declaredSize)
                throw new InvalidDataException("Compound File stream ended before the declared size.");

            byte[] result = new byte[(int)outputLength];
            int written = 0;

            for (int i = 0; i < sectors.Count && written < result.Length; i++)
            {
                int copy = Math.Min(sectors[i].Length, result.Length - written);
                Buffer.BlockCopy(sectors[i], 0, result, written, copy);
                written += copy;
            }

            return result;
        }

        private byte[] ReadSector(uint sectorId)
        {
            long offset = ((long)sectorId + 1L) * sectorSize;
            if (offset < sectorSize || offset + sectorSize > stream.Length)
                throw new InvalidDataException("Compound File sector is outside the file.");
            return ReadExactAt(offset, sectorSize);
        }

        private byte[] ReadExactAt(long offset, int count)
        {
            byte[] buffer = new byte[count];
            stream.Position = offset;
            int read = 0;

            while (read < count)
            {
                int current = stream.Read(buffer, read, count - read);
                if (current <= 0)
                    throw new EndOfStreamException("Unexpected end of Compound File.");
                read += current;
            }

            return buffer;
        }

        private static int NormalizeDirectoryId(uint value)
        {
            return value == NoStream || value > int.MaxValue ? -1 : (int)value;
        }

        private static bool IsRegularSector(uint sector)
        {
            return sector != FreeSector &&
                sector != EndOfChain &&
                sector != FatSector &&
                sector != DifatSector &&
                sector < 0xFFFFFFFA;
        }

        private static ushort ReadUInt16(byte[] data, int offset)
        {
            if (data == null || offset < 0 || offset + 2 > data.Length)
                throw new InvalidDataException("Compound File integer is truncated.");
            return (ushort)(data[offset] | (data[offset + 1] << 8));
        }

        private static uint ReadUInt32(byte[] data, int offset)
        {
            if (data == null || offset < 0 || offset + 4 > data.Length)
                throw new InvalidDataException("Compound File integer is truncated.");

            return (uint)(
                data[offset] |
                (data[offset + 1] << 8) |
                (data[offset + 2] << 16) |
                (data[offset + 3] << 24));
        }

        private static ulong ReadUInt64(byte[] data, int offset)
        {
            uint low = ReadUInt32(data, offset);
            uint high = ReadUInt32(data, offset + 4);
            return ((ulong)high << 32) | low;
        }

        public void Dispose()
        {
            if (disposed)
                return;
            disposed = true;
            stream.Dispose();
        }
    }
}
