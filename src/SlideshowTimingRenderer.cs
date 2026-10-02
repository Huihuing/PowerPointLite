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

        public sealed class SlideShowSettingsSpec
        {
            public bool Loop;
            public bool UseTimings = true;
        }

        public static SlideShowSettingsSpec ReadSlideShowSettings(
            string file)
        {
            SlideShowSettingsSpec result =
                new SlideShowSettingsSpec();

            using (ZipArchive zip = ZipFile.OpenRead(file))
            {
                XmlDocument presentation =
                    LoadXml(zip, "ppt/presentation.xml");

                XmlNode showPr =
                    FindFirst(presentation, "showPr");

                if (showPr == null)
                    return result;

                string loop =
                    GetAttr(showPr, "loop");

                result.Loop =
                    loop == "1" ||
                    string.Equals(
                        loop,
                        "true",
                        StringComparison.OrdinalIgnoreCase);

                string useTimings =
                    GetAttr(showPr, "useTimings");

                if (!string.IsNullOrEmpty(useTimings))
                {
                    result.UseTimings =
                        useTimings != "0" &&
                        !string.Equals(
                            useTimings,
                            "false",
                            StringComparison.OrdinalIgnoreCase);
                }
            }

            return result;
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
