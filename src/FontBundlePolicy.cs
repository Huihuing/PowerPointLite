using System;
using System.Collections.Generic;
using System.IO;

namespace PptxViewer
{
    internal sealed class BundledFontLicenseEntry
    {
        public string FamilyName { get; set; }
        public string FontFilePath { get; set; }
        public string LicenseFilePath { get; set; }
        public string LicenseName { get; set; }
        public bool RedistributionConfirmed { get; set; }
        public bool ApplicationEmbeddingConfirmed { get; set; }
        public string ReviewNote { get; set; }

        public BundledFontLicenseEntry()
        {
            FamilyName = string.Empty;
            FontFilePath = string.Empty;
            LicenseFilePath = string.Empty;
            LicenseName = string.Empty;
            ReviewNote = string.Empty;
        }
    }

    internal sealed class FontBundleDecision
    {
        public bool Allowed { get; set; }
        public string Reason { get; set; }
        public FontEmbeddingDecision MetadataDecision { get; set; }

        public FontBundleDecision()
        {
            Reason = string.Empty;
        }
    }

    internal static class FontBundlePolicy
    {
        public static FontBundleDecision Evaluate(
            BundledFontLicenseEntry entry)
        {
            FontBundleDecision result = new FontBundleDecision();

            if (entry == null)
            {
                result.Reason = "No bundle license entry was provided.";
                return result;
            }

            if (string.IsNullOrEmpty(entry.FontFilePath) ||
                !File.Exists(entry.FontFilePath))
            {
                result.Reason = "The font file is missing.";
                return result;
            }

            if (string.IsNullOrEmpty(entry.LicenseFilePath) ||
                !File.Exists(entry.LicenseFilePath))
            {
                result.Reason = "A local copy of the reviewed font license is required before bundling.";
                return result;
            }

            if (!entry.RedistributionConfirmed)
            {
                result.Reason = "Font file redistribution has not been explicitly confirmed from the license text.";
                return result;
            }

            if (!entry.ApplicationEmbeddingConfirmed)
            {
                result.Reason = "Application embedding/bundling permission has not been explicitly confirmed.";
                return result;
            }

            result.MetadataDecision = FontLicenseService.EvaluateEmbedding(
                entry.FontFilePath,
                FontEmbeddingPurpose.ApplicationBundle,
                true);

            if (result.MetadataDecision == null ||
                !result.MetadataDecision.Allowed)
            {
                result.Reason = result.MetadataDecision == null
                    ? "Font bundle metadata evaluation failed."
                    : result.MetadataDecision.Reason;
                return result;
            }

            result.Allowed = true;
            result.Reason =
                "Bundling is allowed only for this explicitly reviewed font entry. " +
                "Keep the license file and THIRD_PARTY_NOTICES entry with the distributed application.";
            return result;
        }
    }

    internal sealed class BundledFontRegistry
    {
        private readonly List<BundledFontLicenseEntry> entries =
            new List<BundledFontLicenseEntry>();

        public IList<BundledFontLicenseEntry> Entries
        {
            get { return entries.AsReadOnly(); }
        }

        public void AddReviewed(BundledFontLicenseEntry entry)
        {
            FontBundleDecision decision = FontBundlePolicy.Evaluate(entry);

            if (!decision.Allowed)
            {
                throw new InvalidOperationException(
                    "Font cannot be added to the bundle allow-list: " +
                    decision.Reason);
            }

            entries.Add(entry);
        }

        public bool ContainsFamily(string familyName)
        {
            if (string.IsNullOrEmpty(familyName))
                return false;

            for (int i = 0; i < entries.Count; i++)
            {
                if (string.Equals(
                        entries[i].FamilyName,
                        familyName,
                        StringComparison.CurrentCultureIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
