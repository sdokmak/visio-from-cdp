// Toggle Debug mode here
//#define DEBUG

using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.IO.Compression;
using System;
using System.IO.Packaging;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

//C:.
//├───docProps
//├───visio
//│   ├───embeddings
//│   │   └───_rels
//│   ├───masters
//│   │   └───_rels
//│   ├───media
//│   ├───pages
//│   │   └───_rels
//│   └───_rels
//└───_rels

// About Template Digram v0.2.vsdx
// page1.xml: Background-1 this is in fId element
// page2.xml: Revision Control
// page3.xml: Symbols
// page4.xml: Page 1


// Need to complete the following in the branch:
// 1. Read the Json file.
//      - "1.3.6.1.2.1.47.1.1.1.1.13.x": Model
//      - "1.3.6.1.2.1.1.1.0": Description

namespace VisioAutomation {
    class Backend {
        public static Package OpenPackage(string outputPath) {
            Package visioPackage = null;
            DirectoryInfo dirInfo = new DirectoryInfo(Directory.GetParent(outputPath).ToString());
            // Get the Visio file from the location.
            FileInfo[] fileInfos = dirInfo.GetFiles(Regex.Match(outputPath, "[^\\\\]*$").Value);
            if (fileInfos.Count() > 0) {
                FileInfo fileInfo = fileInfos[0];
                string filePathName = fileInfo.FullName;
                // Open the Visio file as a package with read/write file access.
                visioPackage = Package.Open(
                    filePathName,
                    FileMode.Open,
                    FileAccess.ReadWrite);
            }
            // Return the Visio file as a package.
            return visioPackage;
        }
        public static void IteratePackageParts(Package filePackage) {
            // Get all of the package parts contained in the package
            // and then write the URI and content type of each one to the console.
            PackagePartCollection packageParts = filePackage.GetParts();
            foreach (PackagePart part in packageParts) {
                Console.WriteLine("Package part URI: {0}", part.Uri);
                Console.WriteLine("Content type: {0}", part.ContentType.ToString());
            }
        }
        public static PackagePart GetPackagePart(Package filePackage, string relationship) {
            // Use the namespace that describes the relationship 
            // to get the relationship.
            PackageRelationship packageRel = filePackage.GetRelationshipsByType(relationship).FirstOrDefault();
            PackagePart part = null;
            // If the Visio file package contains this type of relationship with 
            // one of its parts, return that part.
            if (packageRel != null) {
                // Clean up the URI using a helper class and then get the part.
                Uri docUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), packageRel.TargetUri);
                part = filePackage.GetPart(docUri);
            }
            return part;
        }
        public static PackagePart GetPackagePart(Package filePackage, PackagePart sourcePart, string relationship) {
            // This gets only the first PackagePart that shares the relationship
            // with the PackagePart passed in as an argument. You can modify the code
            // here to return a different PackageRelationship from the collection.
            PackageRelationship packageRel = sourcePart.GetRelationshipsByType(relationship).FirstOrDefault();
            //PackageRelationship packageRel = sourcePart.GetRelationshipsByType(relationship).ElementAt(133);
            //PackageRelationship packageRel = sourcePart.GetRelationship();
            PackagePart relatedPart = null;
            if (packageRel != null) {
                // Use the PackUriHelper class to determine the URI of PackagePart
                // that has the specified relationship to the PackagePart passed in
                // as an argument.
                Uri partUri = PackUriHelper.ResolvePartUri(sourcePart.Uri, packageRel.TargetUri);
                relatedPart = filePackage.GetPart(partUri);
            }
            return relatedPart;
        }
        public static XDocument GetXMLFromPart(PackagePart packagePart) {
            XDocument partXml = null;
            // Open the packagePart as a stream and then 
            // open the stream in an XDocument object.
            Stream partStream = packagePart.GetStream();
            partXml = XDocument.Load(partStream);
            return partXml;
        }
        public static IEnumerable<XElement> GetXElementsByName(XDocument packagePart, string elementType) {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements = from element in packagePart.Descendants() where element.Name.LocalName == elementType select element;
            // Return the selected elements to the calling code.
            return elements.DefaultIfEmpty(null);
        }
        public static XElement GetXElementByAttribute(IEnumerable<XElement> elements, string attributeName, string attributeValue) {
            // Construct a LINQ query that selects elements from a group of elements by the value of a specific attribute.
            IEnumerable<XElement> selectedElements = from el in elements where el.Attribute(attributeName).Value == attributeValue select el;
            // If there aren't any elements of the specified type with the specified attribute value in the document, return null to the calling code.
            return selectedElements.DefaultIfEmpty(null).FirstOrDefault();
        }
        public static IEnumerable<XElement> GetXElementsByAttribute(IEnumerable<XElement> elements, string attributeName, string attributeValue) {
            // Construct a LINQ query that selects elements from a group of elements by the value of a specific attribute.
            IEnumerable<XElement> selectedElements = from el in elements where el.Attribute(attributeName).Value == attributeValue select el;
            // If there aren't any elements of the specified type with the specified attribute value in the document, return null to the calling code.
            return selectedElements.DefaultIfEmpty(null);
        }
        public static void SaveXDocumentToPart(PackagePart packagePart, XDocument partXML) {
            // Create a new XmlWriterSettings object to define the characteristics for the XmlWriter
            XmlWriterSettings partWriterSettings = new XmlWriterSettings();
            partWriterSettings.Encoding = Encoding.UTF8;
            // Create a new XmlWriter and then write the XML back to the document part.
            XmlWriter partWriter = XmlWriter.Create(packagePart.GetStream(),
                partWriterSettings);
            partXML.WriteTo(partWriter);
            // Flush and close the XmlWriter.
            partWriter.Flush();
            partWriter.Close();
        }
        public static int CheckForRecalc(XDocument customPropsXDoc) {
            // Set the inital pidValue to -1, which is not an allowed value. The calling code tests to see whether the pidValue is greater than - 1.
            int pidValue = -1;
            // Get all of the property elements from the document. 
            IEnumerable<XElement> props = GetXElementsByName(customPropsXDoc, "property");
            // Get the RecalcDocument property from the document if it exists already.
            XElement recalcProp = GetXElementByAttribute(props, "name", "RecalcDocument");
            // If there is already a RecalcDocument instruction in the Custom File Properties part, then we don't need to add another one. Otherwise, we need to create a unique pid value.
            if (recalcProp != null) {
                return pidValue;
            }
            else {
                // Get all of the pid values of the property elements and then convert the IEnumerable object into an array.
                IEnumerable<string> propIDs = from prop in props where prop.Name.LocalName == "property" select prop.Attribute("pid").Value;
                string[] propIDArray = propIDs.ToArray();
                // Increment this id value until a unique value is found. This starts at 2, because 0 and 1 are not valid pid values.
                int id = 2;
                while (pidValue == -1) {
                    if (propIDArray.Contains(id.ToString())) {
                        id++;
                    }
                    else {
                        pidValue = id;
                    }
                }
            }
            return pidValue;
        }
        public static void CreateAndModifyRelsFiles(Package filePackage, PackagePart filePackagePart, XDocument partXML, Uri packageLocation, string contentType, string relationship) { // add iterator
            PackagePart fileRelationPart;
            // Need to check first to see whether the part exists already.
            if (!filePackage.PartExists(packageLocation)) {
                // Create a new blank package part at the specified URI of the specified content type.
                PackagePart newPackagePart = filePackage.CreatePart(packageLocation, contentType);
                XDocument relXML = VisioData.PageRelXML();
                PackagePart newRelationPart = filePackage.CreatePart(new Uri("/visio/pages/_rels/page5.xml.rels", UriKind.Relative), "application/vnd.openxmlformats-package.relationships+xml"); // iterate uri++
                fileRelationPart = newPackagePart;
                // Create a stream from the package part and save the XML document to the package part.
                using (Stream partStream = newPackagePart.GetStream(FileMode.Create, FileAccess.ReadWrite)) {
                    partXML.Save(partStream);
                }
                using (Stream partStream = newRelationPart.GetStream(FileMode.Create, FileAccess.ReadWrite)) {
                    relXML.Save(partStream);
                }
                fileRelationPart.CreateRelationship(new Uri("page1.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/page", "rId1");
            }
            // Create Relationship file content

            // Add a relationship from the file package to this package part. You can also create relationships between an existing package part and a new part.
            // This goes into pages.xml.rels
            filePackagePart.CreateRelationship(new Uri("page5.xml", UriKind.Relative), TargetMode.Internal, relationship, "rId5"); // iterate rId++
        }
        public static XDocument DuplicatePage(XElement page) {
            XNamespace pagens = "http://schemas.microsoft.com/office/visio/2012/main";

            XElement newpage = new XElement(page);
            newpage.FirstAttribute.Value = "9";
            newpage.Attribute("Name").Value = "Page 2";
            //XNamespace pagens = "http://schemas.microsoft.com/visio/2010/relationships/page";
            XDocument newPageDoc = new XDocument(new XElement(newpage));
            return newPageDoc;
        }
        public static XDocument AddToPages(XDocument doc, XElement page) {
            // old code
            XNamespace pagens = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement newpage = new XElement(page);
            newpage.FirstAttribute.Value = "9";
            newpage.Attribute("Name").Value = "Page 2";
            newpage.Attribute("NameU").Value = "Generated Page";
            newpage.Elements().ElementAt(1).Attributes().ElementAt(0).Value = "rId5";
            doc.Root.Add(newpage);
            return doc;
        }
        public static void RecalcDocument(Package filePackage) {
            // Get the Custom File Properties part from the package and then extract the XML from it.
            PackagePart customPart = GetPackagePart(filePackage, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + "custom-properties");
            XDocument customPartXML = GetXMLFromPart(customPart);
            // Check to see whether document recalculation has already been set for this document. If it hasn't, use the integer value returned by CheckForRecalc as the property ID.
            int pidValue = CheckForRecalc(customPartXML);
            if (pidValue > -1) {
                XElement customPartRoot = customPartXML.Elements().ElementAt(0);
                // Two XML namespaces are needed to add XML data to this document. Here, we're using the GetNamespaceOfPrefix and GetDefaultNamespace methods to get the namespaces that 
                // we need. You can specify the exact strings for the namespaces, but that is not recommended.
                XNamespace customVTypesNS = customPartRoot.GetNamespaceOfPrefix("vt");
                XNamespace customPropsSchemaNS = customPartRoot.GetDefaultNamespace();
                // Construct the XML for the new property in the XDocument.Add method. 
                // This ensures that the XNamespace objects will resolve properly, apply the correct prefix, and will not default to an empty namespace.
                customPartRoot.Add(
                    new XElement(customPropsSchemaNS + "property",
                        new XAttribute("pid", pidValue.ToString()),
                        new XAttribute("name", "RecalcDocument"),
                        new XAttribute("fmtid",
                            "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                        new XElement(customVTypesNS + "bool", "true")
                    ));
            }
            // Save the Custom Properties package part back to the package.
            SaveXDocumentToPart(customPart, customPartXML);
        }
        public static void ConvertTemplateToDoc(string templatePath, string outputPath) {
            //File.Copy(directory + "\\" + "NTT Visio Template - A3-Landscape.vstx", directory + "\\" + "Generated Devices.vsdx", true);
            File.Copy(templatePath, outputPath, true);
            using (var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update)) {
                var entry = archive.GetEntry("[Content_Types].xml");
                StringBuilder sb = new StringBuilder();
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append("<Types xmlns =\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.oleObject\"/>" +
                    "<Default Extension=\"bmp\" ContentType=\"image/bmp\"/>" +
                    "<Default Extension=\"emf\" ContentType=\"image/x-emf\"/>" +
                    "<Default Extension=\"png\" ContentType=\"image/png\"/>" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/visio/document.xml\" ContentType=\"application/vnd.ms-visio.drawing.main+xml\"/>" +
                    "<Override PartName=\"/visio/masters/masters.xml\" ContentType=\"application/vnd.ms-visio.masters+xml\"/>");
                for (int i = 1; i <= 19+VisioData.DeviceTypes.Length; i++) {
                    sb.Append($"<Override PartName=\"/visio/masters/master{i}.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/>");
                }
                    sb.Append("<Override PartName=\"/visio/pages/pages.xml\" ContentType=\"application/vnd.ms-visio.pages+xml\"/>" +
                    "<Override PartName=\"/visio/pages/page1.xml\" ContentType=\"application/vnd.ms-visio.page+xml\"/>" +
                    "<Override PartName=\"/visio/pages/page2.xml\" ContentType=\"application/vnd.ms-visio.page+xml\"/>" +
                    "<Override PartName=\"/visio/windows.xml\" ContentType=\"application/vnd.ms-visio.windows+xml\"/>" +
                    "<Override PartName=\"/visio/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" +
                    "<Override PartName=\"/visio/theme/theme2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" +
                    "<Override PartName=\"/visio/theme/theme3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" +
                    "<Override PartName=\"/visio/theme/theme4.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>" +
                    "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" +
                    "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
                    "<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>" +
                    "</Types>");
                entry.Delete();
                entry = archive.CreateEntry("[Content_Types].xml");
                using StreamWriter writer = new StreamWriter(entry.Open());
                writer.Write(sb);
            }
        }
        public static void addMasterstoContentType(string outputPath) {
            using (var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update)) {
                StringBuilder sb = new StringBuilder();
                var entry = archive.GetEntry("[Content_Types].xml");
                sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sb.Append("<Types xmlns =\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.oleObject\"/><Default Extension=\"bmp\" ContentType=\"image/bmp\"/><Default Extension=\"emf\" ContentType=\"image/x-emf\"/><Default Extension=\"png\" ContentType=\"image/png\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/visio/document.xml\" ContentType=\"application/vnd.ms-visio.drawing.main+xml\"/><Override PartName=\"/visio/masters/masters.xml\" ContentType=\"application/vnd.ms-visio.masters+xml\"/><Override PartName=\"/visio/masters/master1.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master2.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master3.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master4.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master5.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master6.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master7.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master8.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master9.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master10.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master11.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master12.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master13.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master14.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master15.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master16.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master17.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master18.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/masters/master19.xml\" ContentType=\"application/vnd.ms-visio.master+xml\"/><Override PartName=\"/visio/pages/pages.xml\" ContentType=\"application/vnd.ms-visio.pages+xml\"/><Override PartName=\"/visio/pages/page1.xml\" ContentType=\"application/vnd.ms-visio.page+xml\"/><Override PartName=\"/visio/pages/page2.xml\" ContentType=\"application/vnd.ms-visio.page+xml\"/><Override PartName=\"/visio/windows.xml\" ContentType=\"application/vnd.ms-visio.windows+xml\"/><Override PartName=\"/visio/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/visio/theme/theme2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/visio/theme/theme3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/visio/theme/theme4.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/><Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/><Override ContentType=\"application/vnd.ms-visio.master+xml\" PartName=\"/visio/masters/master20.xml\"/><Override ContentType=\"application/vnd.ms-visio.master+xml\" PartName=\"/visio/masters/master21.xml\"/>Override ContentType=\"application/vnd.ms-visio.master+xml\" PartName=\"/visio/masters/master22.xml\"/><Override ContentType=\"application/vnd.ms-visio.master+xml\" PartName=\"/visio/masters/master23.xml\"/></Types>");
                entry.Delete();
                entry = archive.CreateEntry("[Content_Types].xml");
                using StreamWriter writer = new StreamWriter(entry.Open());
                writer.Write(sb);
                writer.Close();
            }
        }
        public static int[] Initarray(Array array, int startingValue) {
            int[] initialisedArray = new int[array.Length];
            for (int i = 0; i < array.Length; i++) {
                initialisedArray[i] = startingValue;
            }
            return initialisedArray;
        }
        public static void Swap(IList<int> list, int indexA, int indexB) {
            int tmp = list[indexA];
            list[indexA] = list[indexB];
            list[indexB] = tmp;
        }
    }
}

// Archive:
//CreateNewPackagePart1(visioPackage, pagePart, relationXML, new Uri("/visio/pages/_rels/page5.xml", UriKind.Relative), "application/vnd.openxmlformats-package.relationships+xml", "http://schemas.microsoft.com/visio/2010/relationships/page");
// Get all of the shapes from the page by getting
// all of the Shape elements from the pageXML document.
//IEnumerable<XElement> shapesXML = GetXElementsByName(pageXML, "Shape");


// Select a Shape element from the shapes on the page by 
// its name. You can modify this code to select elements
// by other attributes and their values.
//XElement startEndShapeXML = GetXElementByAttribute(shapesXML, "NameU", "Start/End");


// Query the XML for the shape to get the Text element, and
// return the first Text element node.
//IEnumerable<XElement> textElements = from element in startEndShapeXML.Elements() where element.Name.LocalName == "Text" select element;
//XElement textElement = textElements.ElementAt(0);



// Change the shape text, leaving the <cp> element alone.
//textElement.LastNode.ReplaceWith("End process");
//----------------------------------------------------
// Backend.IteratePackageParts(visioPackage);
// Console.WriteLine("Press any key to continue ...");
// Console.ReadKey();