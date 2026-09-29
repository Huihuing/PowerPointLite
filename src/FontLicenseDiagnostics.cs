using System;
using System.IO;

namespace PptxViewer
{
    internal static class FontLicenseDiagnostics
    {
        public static void ValidatePolicyAndParser()
        {
            string root = Path.Combine(
                Path.GetTempPath(),
                "PowerPointLite_FontLicenseSelfTest_" +
                Guid.NewGuid().ToString("N"));

            Directory.CreateDirectory(root);

            try
            {
                ValidateCase(
                    root,
                    "installable.ttf",
                    0x0000,
                    FontEmbeddingLevel.Installable,
                    true,
                    true,
                    true);

                ValidateCase(
                    root,
                    "restricted.ttf",
                    0x0002,
                    FontEmbeddingLevel.Restricted,
                    false,
                    false,
                    true);

                ValidateCase(
                    root,
                    "preview-print.ttf",
                    0x0004,
                    FontEmbeddingLevel.PreviewAndPrint,
                    true,
                    false,
                    true);

                ValidateCase(
                    root,
                    "editable-no-subset.ttf",
                    0x0108,
                    FontEmbeddingLevel.Editable,
                    true,
                    true,
                    false);

                ValidateCase(
                    root,
                    "bitmap-only.ttf",
                    0x0208,
                    FontEmbeddingLevel.Unknown,
                    false,
                    false,
                    true);

                string installable = Path.Combine(root, "installable.ttf");

                FontEmbeddingDecision denied =
                    FontLicenseService.EvaluateEmbedding(
                        installable,
                        FontEmbeddingPurpose.EditableDocument,
                        false);

                if (denied == null || denied.Allowed)
                {
                    throw new InvalidOperationException(
                        "Embedding must remain denied when the license text has not been explicitly reviewed.");
                }

                FontEmbeddingDecision allowed =
                    FontLicenseService.EvaluateEmbedding(
                        installable,
                        FontEmbeddingPurpose.EditableDocument,
                        true);

                if (allowed == null || !allowed.Allowed)
                {
                    throw new InvalidOperationException(
                        "Installable fsType plus an explicit license confirmation should allow editable-document embedding.");
                }

                FontEmbeddingDecision bundleDenied =
                    FontLicenseService.EvaluateEmbedding(
                        installable,
                        FontEmbeddingPurpose.ApplicationBundle,
                        false);

                if (bundleDenied == null || bundleDenied.Allowed)
                {
                    throw new InvalidOperationException(
                        "Application font bundling must not be authorized by fsType alone.");
                }

                FontEmbeddingDecision bundleConfirmed =
                    FontLicenseService.EvaluateEmbedding(
                        installable,
                        FontEmbeddingPurpose.ApplicationBundle,
                        true);

                if (bundleConfirmed == null || !bundleConfirmed.Allowed)
                {
                    throw new InvalidOperationException(
                        "Application bundling should require and honor an explicit redistribution/app-embedding license confirmation.");
                }
            }
            finally
            {
                try
                {
                    if (Directory.Exists(root))
                        Directory.Delete(root, true);
                }
                catch { }
            }
        }

        private static void ValidateCase(
            string root,
            string fileName,
            ushort fsType,
            FontEmbeddingLevel expectedLevel,
            bool expectedCanEmbed,
            bool expectedCanEdit,
            bool expectedCanSubset)
        {
            string path = Path.Combine(root, fileName);
            File.WriteAllBytes(path, BuildSyntheticSfnt(fsType));

            FontLicenseInfo info = OpenTypeFontLicenseReader.Read(path);

            if (info == null || !info.MetadataAvailable)
                throw new InvalidOperationException("Synthetic font metadata was not parsed: " + fileName);

            if (info.RawFsType != fsType)
                throw new InvalidOperationException("fsType mismatch for " + fileName);

            if (info.EmbeddingLevel != expectedLevel)
                throw new InvalidOperationException("Embedding level mismatch for " + fileName);

            if (info.CanEmbed != expectedCanEmbed)
                throw new InvalidOperationException("CanEmbed mismatch for " + fileName);

            if (info.CanEmbedForEditing != expectedCanEdit)
                throw new InvalidOperationException("CanEmbedForEditing mismatch for " + fileName);

            if (info.CanSubset != expectedCanSubset)
                throw new InvalidOperationException("CanSubset mismatch for " + fileName);

            if (!info.RequiresLicenseTextReview)
            {
                throw new InvalidOperationException(
                    "The parser must never treat fsType as a substitute for reviewing the actual font license.");
            }
        }

        private static byte[] BuildSyntheticSfnt(ushort fsType)
        {
            byte[] data = new byte[38];

            WriteUInt32BigEndian(data, 0, 0x00010000);
            WriteUInt16BigEndian(data, 4, 1);
            WriteUInt16BigEndian(data, 6, 0);
            WriteUInt16BigEndian(data, 8, 0);
            WriteUInt16BigEndian(data, 10, 0);

            data[12] = (byte)'O';
            data[13] = (byte)'S';
            data[14] = (byte)'/';
            data[15] = (byte)'2';
            WriteUInt32BigEndian(data, 16, 0);
            WriteUInt32BigEndian(data, 20, 28);
            WriteUInt32BigEndian(data, 24, 10);

            WriteUInt16BigEndian(data, 28, 0);
            WriteUInt16BigEndian(data, 30, 0);
            WriteUInt16BigEndian(data, 32, 400);
            WriteUInt16BigEndian(data, 34, 5);
            WriteUInt16BigEndian(data, 36, fsType);

            return data;
        }

        private static void WriteUInt16BigEndian(byte[] data, int offset, ushort value)
        {
            data[offset] = (byte)(value >> 8);
            data[offset + 1] = (byte)value;
        }

        private static void WriteUInt32BigEndian(byte[] data, int offset, uint value)
        {
            data[offset] = (byte)(value >> 24);
            data[offset + 1] = (byte)(value >> 16);
            data[offset + 2] = (byte)(value >> 8);
            data[offset + 3] = (byte)value;
        }
    }
}
