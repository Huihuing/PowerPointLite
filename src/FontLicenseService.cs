using System;

namespace PptxViewer
{
    internal enum FontEmbeddingPurpose
    {
        PreviewAndPrintDocument,
        EditableDocument,
        PdfExport,
        ApplicationBundle
    }

    internal sealed class FontEmbeddingDecision
    {
        public FontLicenseInfo Metadata { get; set; }
        public FontEmbeddingPurpose Purpose { get; set; }
        public bool MetadataAllows { get; set; }
        public bool ExplicitLicenseConfirmed { get; set; }
        public bool Allowed { get; set; }
        public string Reason { get; set; }

        public FontEmbeddingDecision()
        {
            Reason = string.Empty;
        }
    }

    internal static class FontLicenseService
    {
        public static FontEmbeddingDecision EvaluateEmbedding(
            string fontPath,
            FontEmbeddingPurpose purpose,
            bool explicitLicenseConfirmed)
        {
            FontLicenseInfo metadata = OpenTypeFontLicenseReader.Read(fontPath);
            FontEmbeddingDecision decision = new FontEmbeddingDecision();

            decision.Metadata = metadata;
            decision.Purpose = purpose;
            decision.ExplicitLicenseConfirmed = explicitLicenseConfirmed;
            decision.MetadataAllows = MetadataAllowsPurpose(metadata, purpose);

            if (purpose == FontEmbeddingPurpose.ApplicationBundle)
            {
                // OpenType fsType describes document embedding behavior. It does
                // not grant permission to redistribute a TTF/OTF with an app.
                decision.MetadataAllows = false;

                if (!explicitLicenseConfirmed)
                {
                    decision.Allowed = false;
                    decision.Reason = "Application font bundling requires an explicit redistribution/app-embedding license review. fsType alone is not sufficient.";
                    return decision;
                }

                decision.Allowed = true;
                decision.Reason = "Application bundling was explicitly confirmed from the font license. Keep the license text in licenses/ and THIRD_PARTY_NOTICES.md.";
                return decision;
            }

            if (!metadata.MetadataAvailable)
            {
                decision.Allowed = false;
                decision.Reason = "Font embedding metadata is unavailable. Embedding is denied until the font and its license are reviewed.";
                return decision;
            }

            if (!decision.MetadataAllows)
            {
                decision.Allowed = false;
                decision.Reason = "The font fsType metadata does not allow this embedding purpose.";
                return decision;
            }

            if (!explicitLicenseConfirmed)
            {
                decision.Allowed = false;
                decision.Reason = "fsType permits the requested embedding mode, but the actual license text has not been confirmed. This project uses a conservative deny-by-default policy.";
                return decision;
            }

            decision.Allowed = true;
            decision.Reason = "Embedding is allowed by the parsed fsType metadata and an explicit license review was confirmed.";
            return decision;
        }

        private static bool MetadataAllowsPurpose(
            FontLicenseInfo metadata,
            FontEmbeddingPurpose purpose)
        {
            if (metadata == null || !metadata.MetadataAvailable)
                return false;

            if (metadata.EmbeddingLevel == FontEmbeddingLevel.Restricted ||
                metadata.BitmapEmbeddingOnly)
            {
                return false;
            }

            if (purpose == FontEmbeddingPurpose.EditableDocument)
                return metadata.CanEmbedForEditing;

            if (purpose == FontEmbeddingPurpose.PreviewAndPrintDocument ||
                purpose == FontEmbeddingPurpose.PdfExport)
            {
                return metadata.CanPreviewAndPrint;
            }

            return false;
        }
    }
}
