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

    class Program
    {
        static void Test1()
        {
            FileInfo fi = new FileInfo("Test1.pptx");
            if (fi.Exists)
                fi.Delete();
            Dictionary<string, string> tagMap = new Dictionary<string, string>()
                    {
                        {"Title", "A Grand Affair"},
                        {"CustomerName", "Eric White"},
                        {"CustomerCity", "Seattle"},
                        {"TotalRevenue", "$3,000,000"},
                        {"ProjectedRevenue", "$5,500,000"},
                        {"AccountRep", "Tai Yi"},
                        {"Years", "5"},
                        {"TopPercent", "15"},
                    };

            byte[] byteArray = File.ReadAllBytes("Template1.pptx");
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (PresentationDocument pDoc =
                    PresentationDocument.Open(mem, true))
                {
                    PmlTemplateProcessor.ProcessPmlTemplate(pDoc, tagLookup =>
                    {
                        if (tagMap.ContainsKey(tagLookup))
                            return tagMap[tagLookup];
                        else
                            return String.Format("!!!! ERROR Tag:{0} not defined !!!!", tagLookup);
                    }
                    );
                }

                using (FileStream fileStream = new FileStream("Test1.pptx",
                    System.IO.FileMode.CreateNew))
                {
                    mem.WriteTo(fileStream);
                }
            }
        }

        static void Test2()
        {
            FileInfo fi = new FileInfo("Test2.pptx");
            if (fi.Exists)
                fi.Delete();

            byte[] byteArray = File.ReadAllBytes("Template2.pptx");
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (PresentationDocument pDoc =
                    PresentationDocument.Open(mem, true))
                {
                    PmlTemplateProcessor.ProcessPmlTemplate(pDoc, tagXml =>
                    {
                        XElement import = XElement.Parse(tagXml);
                        if (import.Attribute("Name") == null)
                        {
                            if (import.Attribute("ContentType").Value == "Text" &&
                                import.Attribute("Id").Value == "1")
                                return "Conforms to Procedure 1202";
                            if (import.Attribute("ContentType").Value == "Text" &&
                                import.Attribute("Id").Value == "2")
                                return "Adjust as necessary";
                            return "";
                        }
                        if (import.Attribute("Name").Value == "PresentationTitle")
                            return "Overview of Equipment Acquisition Process";
                        if (import.Attribute("Name").Value == "Customer")
                            return "Contoso Equipment Inc.";
                        return "";
                    }
                    );
                }

                using (FileStream fileStream = new FileStream("Test2.pptx",
                    System.IO.FileMode.CreateNew))
                {
                    mem.WriteTo(fileStream);
                }
            }
        }

        static void Main(string[] args)
        {
            Test1();
            Test2();
        }
    }
}