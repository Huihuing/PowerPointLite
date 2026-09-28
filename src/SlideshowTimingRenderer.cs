using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Xml;

namespace PptxViewer
{
    internal static partial class InternalPptxRenderer
    {
        public sealed class SlideAdvanceSpec
        {
            public int AutoAdvanceMs = -1;
            public bool AdvanceOnClick = true;
        }

        public static List<SlideAdvanceSpec> ReadSlideAdvanceSpecs(
            string file)
        {
            List<SlideAdvanceSpec> result =
                new List<SlideAdvanceSpec>();

            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                PresentationInfo info =
                    ReadPresentationInfo(zip);

                for (int i = 0;
                     i < info.SlideParts.Count;
                     i++)
                {
                    SlideAdvanceSpec spec =
                        new SlideAdvanceSpec();

                    XmlDocument document =
                        LoadXml(zip, info.SlideParts[i]);

                    XmlNode transition =
                        FindFirst(document, "transition");

                    if (transition != null)
                    {
                        string advClick =
                            GetAttr(transition, "advClick");

                        if (!string.IsNullOrEmpty(advClick))
                        {
                            spec.AdvanceOnClick =
                                advClick != "0" &&
                                !string.Equals(
                                    advClick,
                                    "false",
                                    StringComparison.OrdinalIgnoreCase);
                        }

                        string advTm =
                            GetAttr(transition, "advTm");

                        int autoAdvanceMs;

                        if (!string.IsNullOrEmpty(advTm) &&
                            int.TryParse(
                                advTm,
                                out autoAdvanceMs) &&
                            autoAdvanceMs >= 0)
                        {
                            spec.AutoAdvanceMs =
                                autoAdvanceMs;
                        }
                    }

                    result.Add(spec);
                }
            }

            return result;
        }
    }
}
