using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public static class Extensions
    {
        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument partXDocument = part.Annotation<XDocument>();
            if (partXDocument != null)
                return partXDocument;
            using (Stream partStream = part.GetStream())
            using (XmlReader partXmlReader = XmlReader.Create(partStream))
                partXDocument = XDocument.Load(partXmlReader);
            part.AddAnnotation(partXDocument);
            return partXDocument;
        }

        public static void PutXDocument(this OpenXmlPart part)
        {
            XDocument partXDocument = part.GetXDocument();
            if (partXDocument != null)
            {
                using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
                using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                    partXDocument.Save(partXmlWriter);
            }
        }

        public static void PutXDocument(this OpenXmlPart part, XDocument document)
        {
            using (Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter partXmlWriter = XmlWriter.Create(partStream))
                document.Save(partXmlWriter);
            part.RemoveAnnotations<XDocument>();
            part.AddAnnotation(document);
        }

        public static string ToStringAlignAttributes(this XContainer xContainer)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.OmitXmlDeclaration = true;
            settings.NewLineOnAttributes = true;
            StringBuilder sb = new StringBuilder();
            using (XmlWriter xmlWriter = XmlWriter.Create(sb, settings))
                xContainer.WriteTo(xmlWriter);
            return sb.ToString();
        }

        public static string StringConcatenate(this IEnumerable<string> source)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in source)
                sb.Append(s);
            return sb.ToString();
        }

        public static IEnumerable<TResult> Rollup<TSource, TResult>(
            this IEnumerable<TSource> source,
            TResult seed,
            Func<TSource, TResult, TResult> projection)
        {
            TResult nextSeed = seed;
            foreach (TSource src in source)
            {
                TResult projectedValue = projection(src, nextSeed);
                nextSeed = projectedValue;
                yield return projectedValue;
            }
        }

        public static IEnumerable<IGrouping<TKey, TSource>> GroupAdjacent<TSource, TKey>(
            this IEnumerable<TSource> source,
            Func<TSource, TKey> keySelector)
        {
            TKey last = default(TKey);
            bool haveLast = false;
            List<TSource> list = new List<TSource>();

            foreach (TSource s in source)
            {
                TKey k = keySelector(s);
                if (haveLast)
                {
                    if (!k.Equals(last))
                    {
                        yield return new GroupOfAdjacent<TSource, TKey>(list, last);
                        list = new List<TSource>();
                        list.Add(s);
                        last = k;
                    }
                    else
                    {
                        list.Add(s);
                        last = k;
                    }
                }
                else
                {
                    list.Add(s);
                    last = k;
                    haveLast = true;
                }
            }
            if (haveLast)
                yield return new GroupOfAdjacent<TSource, TKey>(list, last);
        }
    }

    public class GroupOfAdjacent<TSource, TKey> : IEnumerable<TSource>, IGrouping<TKey, TSource>
    {
        public TKey Key { get; set; }
        private List<TSource> GroupList { get; set; }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return ((System.Collections.Generic.IEnumerable<TSource>)this).GetEnumerator();
        }

        System.Collections.Generic.IEnumerator<TSource>
            System.Collections.Generic.IEnumerable<TSource>.GetEnumerator()
        {
            foreach (var s in GroupList)
                yield return s;
        }

        public GroupOfAdjacent(List<TSource> source, TKey key)
        {
            GroupList = source;
            Key = key;
        }
    }

    public static class PmlTemplateProcessor
    {
        private static object SplitIntoSingleCharRuns(XNode node)
        {
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == a + "p")
                {
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e =>
                        {
                            if (e.Name == a + "r")
                            {
                                string text = e.Elements(a + "t")
                                    .Select(t => (string)t)
                                    .StringConcatenate();
                                return (object)text.Select(c => new XElement(a + "r",
                                        e.Elements().Where(z => z.Name != a + "t"),
                                        new XElement(a + "t",
                                            c.ToString()))
                                    );
                            }
                            return new XElement(e.Name,
                                e.Attributes(),
                                e.Nodes().Select(n => SplitIntoSingleCharRuns(n)));
                        }));
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => SplitIntoSingleCharRuns(n)));
            }
            return node;
        }

        private enum CharacterType
        {
            NotInTag,
            OnTagOpen1,
            OnTagOpen2,
            OnTagClose1,
            OnTagClose2,
            InTag,
        };

        private static object ReplaceTagsTransform(XNode node, Func<string, string> projectionFunc)
        {
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == a + "p")
                {
                    var paragraphChildElementArray = element.Elements().Select((e, i) =>
                        new
                        {
                            ParagraphChild = e,
                            Index = i,
                            Character = (string)e.Element(a + "t"),
                        }
                    ).ToArray();

                    // The following projection goes through all characters and marks the
                    // characters that are part of tags.
                    var markedCharacters = paragraphChildElementArray
                        .Rollup(CharacterType.NotInTag, (r, previousState) =>
                        {
                            switch (previousState)
                            {
                                case CharacterType.NotInTag:
                                    if (r.Character == "<" &&
                                        r.Index <= paragraphChildElementArray.Length - 2 &&
                                        paragraphChildElementArray[r.Index + 1].Character == "#")
                                        return CharacterType.OnTagOpen1;
                                    return CharacterType.NotInTag;
                                case CharacterType.OnTagOpen1:
                                    return CharacterType.OnTagOpen2;
                                case CharacterType.OnTagOpen2:
                                    return CharacterType.InTag;
                                case CharacterType.InTag:
                                    if (r.Character == "#" &&
                                        r.Index <= paragraphChildElementArray.Length - 2 &&
                                        paragraphChildElementArray[r.Index + 1].Character == ">")
                                        return CharacterType.OnTagClose1;
                                    return CharacterType.InTag;
                                case CharacterType.OnTagClose1:
                                    return CharacterType.OnTagClose2;
                                case CharacterType.OnTagClose2:
                                    if (r.Character == "<" &&
                                        r.Index <= paragraphChildElementArray.Length - 2 &&
                                        paragraphChildElementArray[r.Index + 1].Character == "#")
                                        return CharacterType.OnTagOpen1;
                                    return CharacterType.NotInTag;
                                default:
                                    return CharacterType.NotInTag;
                            }
                        });

                    // The following projection merges the character type with the run information.
                    var zipped = paragraphChildElementArray.Zip(markedCharacters, (r, type) =>
                        new
                        {
                            ParagraphChild = r.ParagraphChild,
                            Character = r.Character,
                            Index = r.Index,
                            CharacterType = type,
                        });

                    // The following projection adds a tag count so that can distinguish between
                    // characters that are part of adjacent tags.
                    var counted = zipped.Rollup(0, (r, previous) =>
                        r.CharacterType == CharacterType.OnTagOpen1 ? previous + 1 : previous);

                    // The following projection merges the tag count with the character type info.
                    var zipped2 = zipped.Zip(counted, (x, y) =>
                        new
                        {
                            ParagraphChild = x.ParagraphChild,
                            Character = x.Character,
                            Index = x.Index,
                            State = x.CharacterType,
                            Count = y,
                        });

                    // Group all runs that are part of a tag together.  Group all runs not part
                    // of a tag together.
                    var zipped3 = zipped2.GroupAdjacent(r =>
                    {
                        if (r.State == CharacterType.NotInTag)
                            return -1;
                        return r.Count;
                    });

                    // Create a new paragraph to replace the paragraph that contains tags.
                    XElement newParagraph =
                        new XElement(a + "p",
                            zipped3.Select(g =>
                            {
                                if (g.Key == -1)
                                {
                                    var groupedRuns = g.GroupAdjacent(r =>
                                    {
                                        string z = r.ParagraphChild.Name.ToString();
                                        var z2 = r.ParagraphChild.Element(a + "rPr");
                                        if (z2 != null)
                                            z += z2.ToString();
                                        return z;
                                    });
                                    return (object)groupedRuns.Select(g2 =>
                                    {
                                        string text = g2
                                            .Select(z => z.Character)
                                            .StringConcatenate();
                                        XName name = g2.First().ParagraphChild.Name;
                                        return new XElement(name,
                                            g2.First().ParagraphChild.Attributes(),
                                            g2.First().ParagraphChild.Element(a + "rPr"),
                                            name == a + "r" ?
                                            new XElement(a + "t",
                                                text) : null);
                                    });
                                }
                                string s = g.Select(z => z.Character).StringConcatenate();
                                string tagContents = s.Substring(2, s.Length - 4).Trim();
                                string replacement = projectionFunc(tagContents);
                                return new XElement(a + "r",
                                        g.First().ParagraphChild.Element(a + "rPr"),
                                        new XElement(a + "t", replacement));
                            }));

                    return newParagraph;
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => ReplaceTagsTransform(n, projectionFunc)));
            }
            return node;
        }

        public static void ProcessPmlTemplate(PresentationDocument pDoc,
            Func<string, string> projectionFunc)
        {
            PresentationPart presentationPart = pDoc.PresentationPart;
            foreach (var slidePart in presentationPart.SlideParts)
            {
                XDocument slideXDoc = slidePart.GetXDocument();
                XElement root = slideXDoc.Root;
                XElement newRoot = (XElement)SplitIntoSingleCharRuns(root);
                newRoot = (XElement)ReplaceTagsTransform(newRoot, projectionFunc);
                slidePart.PutXDocument(new XDocument(newRoot));
            }
        }
    }
}
