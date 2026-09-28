using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace PptxViewer
{
internal static partial class InternalPptxRenderer
    {
        private static void LoadThemeColors(
            ZipArchive zip,
            Dictionary<string, Color> colors)
        {
            colors["dk1"] = Color.Black;
            colors["lt1"] = Color.White;
            colors["dk2"] = Color.FromArgb(31, 73, 125);
            colors["lt2"] = Color.FromArgb(238, 236, 225);
            colors["accent1"] = Color.FromArgb(79, 129, 189);
            colors["accent2"] = Color.FromArgb(192, 80, 77);
            colors["accent3"] = Color.FromArgb(155, 187, 89);
            colors["accent4"] = Color.FromArgb(128, 100, 162);
            colors["accent5"] = Color.FromArgb(75, 172, 198);
            colors["accent6"] = Color.FromArgb(247, 150, 70);

            XmlDocument theme = LoadXml(zip, "ppt/theme/theme1.xml");
            if (theme == null)
                return;

            XmlNode clrScheme = FindFirst(theme, "clrScheme");
            if (clrScheme == null)
                return;

            foreach (XmlNode child in clrScheme.ChildNodes)
            {
                string key = child.LocalName;
                Color? c = null;

                XmlNode srgb = FindFirst(child, "srgbClr");
                if (srgb != null)
                    c = ParseHexColor(GetAttr(srgb, "val"));

                if (!c.HasValue)
                {
                    XmlNode sys = FindFirst(child, "sysClr");
                    if (sys != null)
                        c = ParseHexColor(GetAttr(sys, "lastClr"));
                }

                if (c.HasValue)
                    colors[key] = c.Value;
            }
        }

        private static Dictionary<string, RelationshipInfo> LoadRelationshipInfos(
            ZipArchive zip,
            string relPart,
            string sourcePart)
        {
            Dictionary<string, RelationshipInfo> result =
                new Dictionary<string, RelationshipInfo>(StringComparer.OrdinalIgnoreCase);

            XmlDocument doc = LoadXml(zip, relPart);
            if (doc == null)
                return result;

            foreach (XmlNode rel in FindAll(doc, "Relationship"))
            {
                string id = GetAttr(rel, "Id");
                string target = GetAttr(rel, "Target");
                string mode = GetAttr(rel, "TargetMode");
                string type = GetAttr(rel, "Type");

                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target))
                    continue;

                RelationshipInfo item = new RelationshipInfo();
                item.Id = id;
                item.Type = type;
                item.Target = target;
                item.TargetMode = mode;
                item.ResolvedTarget = string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase)
                    ? target
                    : ResolvePart(sourcePart, target);

                result[id] = item;
            }

            return result;
        }

        private static Dictionary<string, string> LoadRelationships(
            ZipArchive zip,
            string relPart,
            string sourcePart)
        {
            Dictionary<string, string> result =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            XmlDocument doc = LoadXml(zip, relPart);
            if (doc == null)
                return result;

            foreach (XmlNode rel in FindAll(doc, "Relationship"))
            {
                string id = GetAttr(rel, "Id");
                string target = GetAttr(rel, "Target");
                string mode = GetAttr(rel, "TargetMode");

                if (string.IsNullOrEmpty(id) ||
                    string.IsNullOrEmpty(target) ||
                    string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                result[id] = ResolvePart(sourcePart, target);
            }

            return result;
        }

        private static string RelationshipPart(string sourcePart)
        {
            sourcePart = NormalizePart(sourcePart);

            int slash = sourcePart.LastIndexOf('/');
            string dir = slash >= 0 ? sourcePart.Substring(0, slash + 1) : "";
            string file = slash >= 0 ? sourcePart.Substring(slash + 1) : sourcePart;

            return dir + "_rels/" + file + ".rels";
        }

        private static string ResolvePart(string sourcePart, string target)
        {
            sourcePart = NormalizePart(sourcePart);
            target = target.Replace('\\', '/');

            if (target.StartsWith("/"))
                return target.TrimStart('/');

            int slash = sourcePart.LastIndexOf('/');
            string baseDir = slash >= 0 ? sourcePart.Substring(0, slash + 1) : "";

            string combined = baseDir + target;
            string[] pieces = combined.Split('/');

            List<string> stack = new List<string>();

            foreach (string piece in pieces)
            {
                if (string.IsNullOrEmpty(piece) || piece == ".")
                    continue;

                if (piece == "..")
                {
                    if (stack.Count > 0)
                        stack.RemoveAt(stack.Count - 1);
                }
                else
                {
                    stack.Add(piece);
                }
            }

            return string.Join("/", stack.ToArray());
        }

        private static string NormalizePart(string part)
        {
            return (part ?? "").Replace('\\', '/').TrimStart('/');
        }

        private static XmlDocument LoadXml(ZipArchive zip, string part)
        {
            part = NormalizePart(part);
            ZipArchiveEntry entry = zip.GetEntry(part);

            if (entry == null)
                return null;

            using (Stream s = entry.Open())
            {
                XmlDocument doc = new XmlDocument();
                doc.PreserveWhitespace = false;
                doc.Load(s);
                return doc;
            }
        }

        private static XmlNode FindFirst(XmlNode node, string localName)
        {
            if (node == null)
                return null;

            if (node.LocalName == localName)
                return node;

            foreach (XmlNode child in node.ChildNodes)
            {
                XmlNode found = FindFirst(child, localName);
                if (found != null)
                    return found;
            }

            return null;
        }

        private static List<XmlNode> FindAll(XmlNode node, string localName)
        {
            List<XmlNode> result = new List<XmlNode>();
            FindAllRecursive(node, localName, result);
            return result;
        }

        private static void FindAllRecursive(
            XmlNode node,
            string localName,
            List<XmlNode> result)
        {
            if (node == null)
                return;

            if (node.LocalName == localName)
                result.Add(node);

            foreach (XmlNode child in node.ChildNodes)
                FindAllRecursive(child, localName, result);
        }

        private static XmlNode DirectChild(XmlNode node, string localName)
        {
            if (node == null)
                return null;

            foreach (XmlNode child in node.ChildNodes)
            {
                if (child.LocalName == localName)
                    return child;
            }

            return null;
        }

        private static string GetRelationshipId(XmlNode node)
        {
            if (node == null || node.Attributes == null)
                return null;

            foreach (XmlAttribute a in node.Attributes)
            {
                if (a.LocalName == "id" &&
                    (a.Prefix == "r" ||
                     a.NamespaceURI.IndexOf("/relationships", StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return a.Value;
                }

                if ((a.LocalName == "embed" || a.LocalName == "link") &&
                    (a.Prefix == "r" ||
                     a.NamespaceURI.IndexOf("/relationships", StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return a.Value;
                }
            }

            return null;
        }

        private static string GetAttr(XmlNode node, string name)
        {
            if (node == null || node.Attributes == null)
                return null;

            foreach (XmlAttribute a in node.Attributes)
            {
                if (a.LocalName == name)
                    return a.Value;
            }

            return null;
        }

        private static long GetLong(XmlNode node, string attr, long fallback)
        {
            string value = GetAttr(node, attr);
            long parsed;

            if (!string.IsNullOrEmpty(value) && long.TryParse(value, out parsed))
                return parsed;

            return fallback;
        }
    }
}
