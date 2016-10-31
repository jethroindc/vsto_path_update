using System;
using System.Collections.Generic;
using System.Text;
using System.IO.Compression;
using System.IO;
using Microsoft.Deployment.WindowsInstaller;
using System.Xml;

namespace UpdateWordTemplateCustomAction
{
    public class CustomActions
    {

        private const string CUSTOM_XML = "docProps/custom.xml";

        [CustomAction]
        public static ActionResult UpdateAddonPath(Session session)
        {
            session.Log("Begin UpdateAddonPath");
            CustomActionData data = session.CustomActionData;
            string templatePath = data["TEMPLATE"];
            if (templatePath == null )
            {
                session.Log("missing parameter 'VSTO'");
                return ActionResult.Failure;
            }
            session.Log("using template: " + templatePath);

            String vstoFilePath = data["VSTO"];
            if (vstoFilePath == null)
            {
                session.Log("missing parameter dest");
                return ActionResult.Failure;
            }
            session.Log("using vsto file: " + vstoFilePath);
            using (FileStream templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.ReadWrite))
            {
                using( ZipArchive templateZip = new ZipArchive( templateStream, ZipArchiveMode.Update ))
                {
                    session.Log("loading document from zip");
                    XmlDocument doc = loadCustomXml(templateZip);
                    session.Log("updating assembly path");
                    updateAssemblyPath(session, doc, vstoFilePath);
                    String updatedDoc = doc.OuterXml;
                    session.Log("replacing custom.xml");
                    replaceCustomFile(templateZip, updatedDoc);
                }
            }
            return ActionResult.Success;
        }

        private static void replaceCustomFile(ZipArchive archive, string updatedDoc)
        {
            ZipArchiveEntry origEntry = archive.GetEntry(CUSTOM_XML);
            origEntry.Delete();
            ZipArchiveEntry newEntry = archive.CreateEntry(CUSTOM_XML);
            using (Stream customXmlStream = newEntry.Open())
            {
                StreamWriter writer = new StreamWriter(customXmlStream);
                writer.Write(updatedDoc);
                writer.Close();
            }
        }

        private static XmlDocument loadCustomXml(ZipArchive archive)
        {
            ZipArchiveEntry entry = archive.GetEntry(CUSTOM_XML);
            String content;
            using (Stream customXmlStream = entry.Open())
            {
                StreamReader customXmlReader = new StreamReader(customXmlStream);
                content = customXmlReader.ReadToEnd();
            }

            // treat this as xml to make life easier
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(content);
            return doc;
        }

        private static void updateAssemblyPath( Session session, XmlDocument doc, string newDest )
        {
            XmlNode assemblyNode = doc.SelectSingleNode(@"//*[name()='property'][@name='_AssemblyLocation']");
            if (assemblyNode == null)
            {
                session.Log("unable to find assemly node!");
            }
            else
            {
                XmlNode valueNode = assemblyNode.FirstChild;
                if (valueNode == null)
                {
                    session.Log("value node is missing!!");
                }
                else
                {
                    valueNode.InnerText = "file:///" + newDest + "|1e52a22b-aeb8-4d47-a0f8-135bb6c087f2|vstolocal";
                }
            }
        }
    }
}
