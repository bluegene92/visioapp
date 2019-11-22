using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ReadVisio
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using (Package fPackage = OpenPackage("*.vsdx", System.Environment.SpecialFolder.Desktop))
                {
                    PackagePartCollection fParts = fPackage.GetParts();
                    foreach (PackagePart fPart in fParts)
                    {
                        Console.WriteLine("Package part: {0}", fPart.Uri);
                        Console.WriteLine("Content type: {0}", fPart.ContentType.ToString());
                    }

                    PackageRelationship packageRelationship = fPackage.GetRelationshipsByType("http://schemas.microsoft.com/visio/2010/relationships/document").FirstOrDefault();



                    PackagePart documentPart = GetPackagePart(fPackage,
    "http://schemas.microsoft.com/visio/2010/relationships/document");
                    PackagePart pagesPart = GetPackagePart(fPackage, documentPart,
    "http://schemas.microsoft.com/visio/2010/relationships/pages");
                    PackagePart pagePart = GetPackagePart(fPackage, pagesPart,
                        "http://schemas.microsoft.com/visio/2010/relationships/page");


                    // Open the XML from the Page Contents part.
                    XDocument pageXML = GetXMLFromPart(pagePart);

                    // Get all of the shapes from the page by getting
                    // all of the Shape elements from the pageXML document.
                    IEnumerable<XElement> shapesXML = GetXElementsByName(pageXML, "Shape");
                    // Select a Shape element from the shapes on the page by 
                    // its name. You can modify this code to select elements
                    // by other attributes and their values.
                    XElement startEndShapeXML =
                        GetXElementByAttribute(shapesXML, "NameU", "Shape2");

                    IEnumerable<XElement> textElements = from element in startEndShapeXML.Elements()
                                                         where element.Name.LocalName == "Text"
                                                         select element;
                    XElement textElement = textElements.ElementAt(0);

                    textElement.LastNode.ReplaceWith("Danny Nguyen");

                    // Save the XML back to the Page Contents part.
                    SaveXDocumentToPart(pagePart, pageXML);
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("Error: {0}", err.Message);
            }
            finally
            {
                Console.Write("\nPress any key to continue ...");
                Console.ReadKey();
            }
        }

        private static Package OpenPackage(string fileName, Environment.SpecialFolder folder)
        {
            Package visioPackage = null;
            string directoryPath = System.Environment.GetFolderPath(
                folder);
            DirectoryInfo dirInfo = new DirectoryInfo(directoryPath);

            FileInfo[] fileInfos = dirInfo.GetFiles(fileName);
            if (fileInfos.Length > 0)
            {
                FileInfo fileInfo = fileInfos[0];
                string filePathName = fileInfo.FullName;
                visioPackage = Package.Open(
                    filePathName,
                    FileMode.Open,
                    FileAccess.ReadWrite);
            }
            return visioPackage;
        }

        private static PackagePart GetPackagePart(Package filePackage,
    string relationship)
        {
            PackageRelationship packageRel =
                filePackage.GetRelationshipsByType(relationship).FirstOrDefault();
            PackagePart part = null;
            if (packageRel != null)
            {
                Uri docUri = PackUriHelper.ResolvePartUri(
                    new Uri("/", UriKind.Relative), packageRel.TargetUri);
                part = filePackage.GetPart(docUri);
            }
            return part;
        }

        private static PackagePart GetPackagePart(Package filePackage,
                        PackagePart sourcePart, string relationship)
        {
            PackageRelationship packageRel =
                sourcePart.GetRelationshipsByType(relationship).FirstOrDefault();
            PackagePart relatedPart = null;
            if (packageRel != null)
            {
                Uri partUri = PackUriHelper.ResolvePartUri(
                    sourcePart.Uri, packageRel.TargetUri);
                relatedPart = filePackage.GetPart(partUri);
            }
            return relatedPart;
        }

        private static XDocument GetXMLFromPart(PackagePart packagePart)
        {
            XDocument partXml = null;
            // Open the packagePart as a stream and then 
            // open the stream in an XDocument object.
            Stream partStream = packagePart.GetStream();
            partXml = XDocument.Load(partStream);
            return partXml;
        }

        private static IEnumerable<XElement> GetXElementsByName(
    XDocument packagePart, string elementType)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                from element in packagePart.Descendants()
                where element.Name.LocalName == elementType
                select element;
            // Return the selected elements to the calling code.
            return elements.DefaultIfEmpty(null);
        }

        private static XElement GetXElementByAttribute(IEnumerable<XElement> elements,
    string attributeName, string attributeValue)
        {
            // Construct a LINQ query that selects elements from a group
            // of elements by the value of a specific attribute.
            IEnumerable<XElement> selectedElements =
                from el in elements
                where el.Attribute(attributeName).Value == attributeValue
                select el;
            // If there aren't any elements of the specified type
            // with the specified attribute value in the document,
            // return null to the calling code.
            return selectedElements.DefaultIfEmpty(null).FirstOrDefault();
        }

        private static void SaveXDocumentToPart(PackagePart packagePart,
    XDocument partXML)
        {

            // Create a new XmlWriterSettings object to 
            // define the characteristics for the XmlWriter
            XmlWriterSettings partWriterSettings = new XmlWriterSettings();
            partWriterSettings.Encoding = Encoding.UTF8;
            // Create a new XmlWriter and then write the XML
            // back to the document part.
            XmlWriter partWriter = XmlWriter.Create(packagePart.GetStream(),
                partWriterSettings);
            partXML.WriteTo(partWriter);
            // Flush and close the XmlWriter.
            partWriter.Flush();
            partWriter.Close();
        }
    }
}
