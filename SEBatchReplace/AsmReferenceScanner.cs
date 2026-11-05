using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SEBatchReplace
{
    public sealed class AsmReferenceMatch
    {
        public string PartName { get; private set; }
        public string EncodingName { get; private set; }
        public IList<long> Offsets { get; private set; }
        public bool Truncated { get; private set; }

        public AsmReferenceMatch(string partName, string encodingName, IList<long> offsets, bool truncated)
        {
            PartName = partName;
            EncodingName = encodingName;
            Offsets = offsets;
            Truncated = truncated;
        }
    }

    public sealed class AsmReferenceScanResult
    {
        public string FileFormat { get; set; }
        public long FileLength { get; set; }
        public IList<AsmReferenceMatch> Matches { get; private set; }

        public AsmReferenceScanResult()
        {
            Matches = new List<AsmReferenceMatch>();
        }
    }

    public static class AsmReferenceScanner
    {
        private static readonly byte[] CompoundFileSignature =
            new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        public static AsmReferenceScanResult Scan(string filePath, IList<string> partNames)
        {
            if (String.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("組立檔路徑不可空白。", "filePath");

            var result = new AsmReferenceScanResult();
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                result.FileLength = stream.Length;
                byte[] header = new byte[8];
                int headerRead = stream.Read(header, 0, header.Length);
                result.FileFormat = headerRead == 8 && BytesEqual(header, CompoundFileSignature)
                    ? "OLE/CFB Compound File"
                    : "非標準 OLE/CFB 簽章（或檔案不足 8 bytes）";

                if (partNames == null)
                    return result;

                foreach (string partName in partNames)
                {
                    if (String.IsNullOrWhiteSpace(partName))
                        continue;

                    string normalized = Path.GetFileName(partName.Trim());
                    byte[] ascii = TryGetAsciiBytes(normalized);
                    if (ascii != null)
                    {
                        IList<long> offsets;
                        bool truncated;
                        FindOffsets(stream, ascii, out offsets, out truncated);
                        if (offsets.Count > 0)
                            result.Matches.Add(new AsmReferenceMatch(normalized, "ASCII", offsets, truncated));
                    }

                    byte[] unicode = Encoding.Unicode.GetBytes(normalized);
                    IList<long> unicodeOffsets;
                    bool unicodeTruncated;
                    FindOffsets(stream, unicode, out unicodeOffsets, out unicodeTruncated);
                    if (unicodeOffsets.Count > 0)
                        result.Matches.Add(new AsmReferenceMatch(normalized, "UTF-16LE", unicodeOffsets, unicodeTruncated));
                }
            }
            return result;
        }

        private static byte[] TryGetAsciiBytes(string value)
        {
            for (int i = 0; i < value.Length; i++)
                if (value[i] > 0x7F)
                    return null;
            return Encoding.ASCII.GetBytes(value);
        }

        private static bool BytesEqual(byte[] left, byte[] right)
        {
            if (left.Length != right.Length)
                return false;
            for (int i = 0; i < left.Length; i++)
                if (left[i] != right[i])
                    return false;
            return true;
        }

        private static void FindOffsets(Stream stream, byte[] pattern, out IList<long> offsets, out bool truncated)
        {
            offsets = new List<long>();
            truncated = false;
            if (pattern == null || pattern.Length == 0)
                return;

            int[] prefix = BuildPrefix(pattern);
            int matched = 0;
            long position = 0;
            byte[] buffer = new byte[65536];
            stream.Position = 0;
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                for (int i = 0; i < read; i++, position++)
                {
                    byte current = FoldAscii(buffer[i]);
                    while (matched > 0 && current != FoldAscii(pattern[matched]))
                        matched = prefix[matched - 1];
                    if (current == FoldAscii(pattern[matched]))
                        matched++;
                    if (matched == pattern.Length)
                    {
                        if (offsets.Count < 10)
                            offsets.Add(position - pattern.Length + 1);
                        else
                            truncated = true;
                        matched = prefix[matched - 1];
                    }
                }
            }
        }

        private static int[] BuildPrefix(byte[] pattern)
        {
            var prefix = new int[pattern.Length];
            int length = 0;
            for (int i = 1; i < pattern.Length; i++)
            {
                byte current = FoldAscii(pattern[i]);
                while (length > 0 && current != FoldAscii(pattern[length]))
                    length = prefix[length - 1];
                if (current == FoldAscii(pattern[length]))
                    length++;
                prefix[i] = length;
            }
            return prefix;
        }

        private static byte FoldAscii(byte value)
        {
            if (value >= (byte)'a' && value <= (byte)'z')
                return (byte)(value - 32);
            return value;
        }
    }
}
