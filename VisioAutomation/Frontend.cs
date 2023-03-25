// Toggle Debug mode here
//#define DEBUG

using System.Xml.Linq;
using System.IO;
using System;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.IO.Compression;
using System.Text;
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
    class Frontend {
        public static void EmptyPages(Package visioPackage, List<string> site) {
            VisioData.uniquesites = site.Distinct().ToArray();
            Array.Sort(VisioData.uniquesites, StringComparer.InvariantCulture);
            int sitetotal = VisioData.uniquesites.Count();
            // Duplicate Page-1 however many times we need
            // 1. Modify pagesPart
            PackagePart documentPart = Backend.GetPackagePart(visioPackage, "http://schemas.microsoft.com/visio/2010/relationships/document");
            Console.WriteLine("Sites: ");
            for (int i = 0; i < sitetotal; i++) {
                PackagePart pagesPart = Backend.GetPackagePart(visioPackage, documentPart, "http://schemas.microsoft.com/visio/2010/relationships/pages");
                XDocument pagesXML = Backend.GetXMLFromPart(pagesPart);
                IEnumerable<XElement> pagenames = Backend.GetXElementsByName(pagesXML, "Page");
                XElement newInPagesXML = Backend.GetXElementByAttribute(pagenames, "Name", "Page-1");
                XElement newpage = new XElement(newInPagesXML);
                newpage.FirstAttribute.Value = $"{10+i}"; // iterate++
                newpage.Attribute("Name").Value = VisioData.uniquesites[i]; // change to site names
                newpage.Attribute("NameU").Value = VisioData.uniquesites[i]; // change to site names
                newpage.Elements().ElementAt(1).Attributes().ElementAt(0).Value = $"rId{3+i}";
                pagesXML.Root.Add(newpage);
                // 3. Add relationship to Pages rels file + 4. Create rels file for page
                PackagePart fileRelationPart;
                // Need to check first to see whether the part exists already.
                if (!visioPackage.PartExists(new Uri($"/visio/pages/page{3+i}.xml", UriKind.Relative))) {
                    // Create a new blank package part at the specified URI of the specified content type.
                    PackagePart newPackagePart = visioPackage.CreatePart(new Uri($"/visio/pages/page{3+i}.xml", UriKind.Relative), "application/vnd.ms-visio.page+xml");
                    XDocument relXML = VisioData.PageRelXML();
                    PackagePart newRelationPart = visioPackage.CreatePart(new Uri($"/visio/pages/_rels/page{3+i}.xml.rels", UriKind.Relative), "application/vnd.openxmlformats-package.relationships+xml"); // iterate uri++
                    fileRelationPart = newPackagePart;
                    // Create a stream from the package part and save the XML document to the package part.
                    using (Stream partStream = newPackagePart.GetStream(FileMode.Create, FileAccess.ReadWrite)) {
                        VisioData.EmptyPage().Save(partStream);
                    }
                    using (Stream partStream = newRelationPart.GetStream(FileMode.Create, FileAccess.ReadWrite)) {
                        relXML.Save(partStream);
                    }
                    fileRelationPart.CreateRelationship(new Uri("page1.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/page", "rId1");
                    // Add in all the stencils that'll be in the pages
                    for (int j = 0; j < VisioData.DeviceTypes.Length; j++) { 
                        fileRelationPart.CreateRelationship(new Uri($"../masters/master{20+j}.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/master", $"rId{5+j}");
                    }
                }
                // Create Relationship file content

                // Add a relationship from the file package to this package part. You can also create relationships between an existing package part and a new part.
                // This goes into pages.xml.rels
                pagesPart.CreateRelationship(new Uri($"page{3+i}.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/page", $"rId{3+i}"); // iterate rId++
                Backend.SaveXDocumentToPart(pagesPart, pagesXML);
                Console.WriteLine(" " + VisioData.uniquesites[i]);
            }
        }
        public static void Add_Devices(Package visioPackage, List<string> hostname, List<string> devicetype, List<string> ipaddress, List<string> model, List<string> site) {
            string[] uniquesites = site.Distinct().ToArray();
            Array.Sort(uniquesites, StringComparer.InvariantCulture);
            PackagePart documentPart = Backend.GetPackagePart(visioPackage, "http://schemas.microsoft.com/visio/2010/relationships/document");
            for (int i = 0; i < uniquesites.Count(); i++) {
                var index = site.Select((value, index) => new { value, index })
                    .Where(x => x.value == uniquesites[i])
                    .Select(x => x.index)
                    .ToList();
                // reorganise device data
                string[] uniquedevices = devicetype.Distinct().ToArray();
                string[] sitehostname = { };
                string[] sitedevicetype = { };
                string[] siteipaddress = { };
                string[] sitemodel = { };
                foreach (string uniquedevice in uniquedevices) {
                    foreach (int siteindex in index) {
                        if (devicetype[siteindex] == uniquedevice) {
                            sitehostname = sitehostname.Concat(new string[] { hostname[siteindex] }).ToArray();
                            sitedevicetype = sitedevicetype.Concat(new string[] { devicetype[siteindex] }).ToArray();
                            if (devicetype[siteindex] == null) {
                                
                                Console.WriteLine("Debug Point");
                            } 
                            siteipaddress = siteipaddress.Concat(new string[] { ipaddress[siteindex] }).ToArray();
                            sitemodel = sitemodel.Concat(new string[] { model[siteindex] }).ToArray();
                        }
                    }
                }
                PackagePart pagesPart = Backend.GetPackagePart(visioPackage, documentPart, "http://schemas.microsoft.com/visio/2010/relationships/pages");
                //XElement newInPagesXML = Backend.GetXElementByAttribute(pagenames, "Name", uniquesites[i]).LastNode;
                //IEnumerable<XElement> textElements = from element in startEndShapeXML.Elements() where element == "Text" select element
                // Can assign the locations and scaling here, based on index
                PackagePart pagePart = VisioData.GetPageByrId(visioPackage, pagesPart, 3 + i);
                XDocument pageXML = Backend.GetXMLFromPart(pagePart);
                pageXML.Root.Add(new XElement("Shapes"));
                //XElement DeviceShapesXML = new XElement("Shapes");
                
                // Following the Powershell Script for organising devices
                int limit;
                double sf;
                switch (index.Count) {
                    case int n when n > 60:
                        limit = 20;
                        sf = 0.5;
                        break;
                    case int n when n > 48:
                        limit = 10;
                        sf = 1;
                        break;
                    default:
                        limit = 8;
                        sf = 1;
                        break;
                }
                (string master, string name, string nameu, string id) = ("", "", "", "");
                (double width, double height, double textsize) = (0, 0, 0.01388*8);
                int rows = (int)Math.Floor(Convert.ToDouble(index.Count()/limit));
                int[] devicesInRow = Enumerable.Repeat(limit, rows+1).ToArray();
                int[] row_index = { };
                for (int r = 0; r <= rows; r++) {
                    row_index = row_index.Concat(Enumerable.Repeat(r, limit)).ToArray();
                }
                devicesInRow[devicesInRow.Length - 1] = index.Count()%devicesInRow.Last();
                
                if (index.Count() < limit) {
                    sf = 1+(3/index.Count());
                }

                double[] pinX = { };
                double[] pinY = { };
                int k = 0;
                for (int j = 0; j < index.Count(); j++) {
                    if (j/limit != k) {
                        k++;
                    }
                    pinX = pinX.Concat(Enumerable.Repeat(8.27 + (sf * ((-0.75*(devicesInRow[k] - 1)) + 1.5*j - 1.5*(limit * row_index[j]))), 1)).ToArray();
                    pinY = pinY.Concat(Enumerable.Repeat(6.76 + (sf * (1.5*row_index[j]-0.75*rows)), 1)).ToArray();
                }
                double previousPinY = pinY.Last();
                Boolean match = false;
                for (int j = 0; j < index.Count(); j++) {
                    for (int n = 0; n < VisioData.MasterNumber.Length; n++) {
                        if (sitedevicetype[j] == VisioData.DeviceTypes[n]) {
                            (master, name, nameu) = (VisioData.MasterNumber[n].ToString(), VisioData.DeviceTypes[n], VisioData.DeviceTypes[n]);
                            (width, height) = (sf*0.03936*VisioData.DeviceWidth[n], sf*0.03936*VisioData.DeviceHeight[n]);
                            match = true;
                            break;
                        }
                    }
                    if (!match) {
                        Console.WriteLine("Failed to assign shape master. Press any key to continue.");
                        Console.ReadKey();
                        return;
                    }
                    //DeviceShapesXML.Element("Shapes").Add(
                    XElement ShapeXML = new XElement("Shape", new XAttribute("Master", master), new XAttribute("Type", "Foreign"), new XAttribute("Name", name), new XAttribute("NameU", nameu), new XAttribute("ID", $"{j+1}"),
                                        new XElement("Cell", new XAttribute("V", pinX[j]), new XAttribute("N", "PinX")),
                                        new XElement("Cell", new XAttribute("V", pinY[j]), new XAttribute("N", "PinY")),
                                        new XElement("Cell", new XAttribute("V", $"{width}"), new XAttribute("N", "Width")),
                                        new XElement("Cell", new XAttribute("V", $"{height}"), new XAttribute("N", "Height")),
                                        new XElement("Section", new XAttribute("N", "Character"),
                                            new XElement("Row", new XAttribute("IX", "0"),
                                                new XElement("Cell", new XAttribute("V", $"{sf*textsize}"), new XAttribute("N", "Size")))),
                                        new XElement("Text",
                                            sitemodel[j] + "\n",
                                            sitehostname[j] + "\n",
                                            siteipaddress[j]));
                    pageXML.Root.Elements().ElementAt(0).Add(ShapeXML);
                }
                //pageXML.Root.Add(DeviceShapesXML);
                Backend.SaveXDocumentToPart(pagePart, pageXML);
            }
        }
        public static void Add_CDP_Devices(Package visioPackage, List<string> hostname, List<string> devicetype, List<string> ipaddress, List<string> model, List<string> site, List<int> tier, List<int> indtier, string[][,] interfaces) {
            string[] uniquesites = site.Distinct().ToArray();
            Array.Sort(uniquesites, StringComparer.InvariantCulture);
            PackagePart documentPart = Backend.GetPackagePart(visioPackage, "http://schemas.microsoft.com/visio/2010/relationships/document");
            for (int i = 0; i < uniquesites.Count(); i++) {
                var index = site.Select((value, index) => new { value, index })
                    .Where(x => x.value == uniquesites[i])
                    .Select(x => x.index)
                    .ToList();

                PackagePart pagesPart = Backend.GetPackagePart(visioPackage, documentPart, "http://schemas.microsoft.com/visio/2010/relationships/pages");
                //XElement newInPagesXML = Backend.GetXElementByAttribute(pagenames, "Name", uniquesites[i]).LastNode;
                //IEnumerable<XElement> textElements = from element in startEndShapeXML.Elements() where element == "Text" select element
                // Can assign the locations and scaling here, based on index
                PackagePart pagePart = VisioData.GetPageByrId(visioPackage, pagesPart, 3 + i);
                XDocument pageXML = Backend.GetXMLFromPart(pagePart);
                pageXML.Root.Add(new XElement("Shapes"));
                //XElement DeviceShapesXML = new XElement("Shapes");

                // Following the Powershell Script for organising devices
                int limit;
                double sf;
                switch (index.Count) {
                    case int n when n > 60:
                        limit = 20;
                        sf = 0.5;
                        break;
                    case int n when n > 48:
                        limit = 10;
                        sf = 1;
                        break;
                    default:
                        limit = 8;
                        sf = 1;
                        break;
                }
                (string master, string name, string nameu, string id) = ("", "", "", "");
                (double width, double height, double textsize) = (0, 0, 0.01388 * 8);

                // we need the sitetier and siteindtier as well.
                // this is all wrong
                // adding tier of index (within site so correct)
                List<int> sitetier = new List<int>();
                List<int> siteindtier = new List<int>();
                for (int j=0; j < index.Count; j++) {
                    sitetier.Add(tier[index[j]]);
                    siteindtier.Add(indtier[index[j]]);
                }

                ////// leverage limit if tier is higher.
                //if (siteindtier.Max() > limit) {
                //    limit = siteindtier.Max();
                //}
                //int rows = (int)Math.Floor(Convert.ToDouble(index.Count() / limit));
                int rows = sitetier.Max();
                // devicesInRow is an int[] with the count of each tier of the same site
                int[] devicesInRow = new int[rows];
                for (int j=0; j < rows; j++) {
                    devicesInRow[j] = sitetier.Where(x => x.Equals(j+1)).Count();
                }
                //int[] devicesInRow = Enumerable.Repeat(limit, rows + 1).ToArray();
                int[] row_index = { }; // this is only used if the siteindtier > limit
                int[] actual_row_index = { }; // tier rows
                int rebase_for_actual_row = 0;
                // add a row to devicesInRow if Limit is reached
                int rebase = 0;
                for (int r = 0; r < rows; r++) {
                    if (devicesInRow[r] > limit) {
                        rebase++;
                    }
                    // an array of however many devices
                    row_index = row_index.Concat(Enumerable.Repeat(rebase, devicesInRow[r])).ToArray();
                    // to reorder based on site tier and siteindtier
                    actual_row_index = actual_row_index.Concat(Enumerable.Repeat(rebase_for_actual_row+rebase, devicesInRow[r])).ToArray();
                    rebase_for_actual_row++;
                }


                List<int> indexd = new List<int>() { 3,1,0,5,0 };

                if (index.Count() < limit) {
                    sf = 1 + (3 / index.Count());
                }
                
                double[] pinX = { };
                double[] pinY = { };
                int k = 0;
                int l = 0;
                int old_k = 0;
                // How to iterate based on 
                // 1. tier, then 
                // 2. indtier
                // let's create another list/array with the order being correct

                for (int j = 0; j < index.Count(); j++) {
                    // need a variable that iterates until reaching max devices in row, but tier and ind tier just don't work for this case...
                    if (l >= devicesInRow[k] && k + 1 < devicesInRow.Length) {
                        k++;
                        l = 0;
                        old_k = k - 1;
                    }
                    else if (l >= devicesInRow[k]) {
                        l = 0;
                    }
                    pinX = pinX.Concat(Enumerable.Repeat(8.27 + (sf * ((-0.75 * (devicesInRow[k] - 1)) + 1.5 * l - 1.5 * (limit * row_index[j]))), 1)).ToArray();
                    pinY = pinY.Concat(Enumerable.Repeat(6.76 - (sf * (1.5 * actual_row_index[j] - 0.75 * (rows-1))), 1)).ToArray();
                    l++;
                }
                double previousPinY = pinY.Last();
                Boolean match = false;

                /// Reorder 'index' based on sitetier and indtier. Right now it's lowest to highest based the index number but it should be 
                /// based on [sitetier(aka row), siteindtier(aka column)]
                List<int> index_reordered = new List<int>();

                for (int o= sitetier.Min(); o <= sitetier.Max(); o++) {
                    // problem is can only get one index, not all... 
                    // Array.IndexOf(sitetier.ToArray(), 2)
                    // while 
                    var o_index = Enumerable.Range(0, sitetier.Count)
                        .Where(i => sitetier[i] == o)
                        .ToList();
                    for (int p=siteindtier.Min(); p <= siteindtier.Max(); p++) {
                        for (int pp=1; pp <= o_index.Count; pp++) { 
                            if (siteindtier[o_index[pp-1]] == p) {
                                index_reordered.Add(index[o_index[pp-1]]);
                                break;
                            }
                        }
                    }
                }
                string stenciltext = "";
                for (int j = 0; j < index_reordered.Count(); j++) {
                    stenciltext = "";
                    for (int n = 0; n < VisioData.MasterNumber.Length; n++) {
                        if (devicetype[index_reordered[j]] == VisioData.DeviceTypes[n]) {
                            (master, name, nameu) = (VisioData.MasterNumber[n].ToString(), VisioData.DeviceTypes[n], VisioData.DeviceTypes[n]);
                            (width, height) = (sf * 0.03936 * VisioData.DeviceWidth[n], sf * 0.03936 * VisioData.DeviceHeight[n]);
                            match = true;
                            break;
                        }
                    }
                    if (!match) {
                        Console.WriteLine("Failed to assign shape master. Press any key to continue.");
                        Console.ReadKey();
                        return;
                    }
                    if (model[index_reordered[j]] != "") {
                        stenciltext += model[index_reordered[j]];
                    }
                    if (hostname[index_reordered[j]] != "") {
                        if (model[index_reordered[j]] != "") {
                            stenciltext += "\n";
                        }
                        stenciltext += hostname[index_reordered[j]];
                    }
                    if (ipaddress[index_reordered[j]] != "") {
                        if (model[index_reordered[j]] != "" || hostname[index_reordered[j]] != "") {
                            stenciltext += "\n";
                        }
                        stenciltext += ipaddress[index_reordered[j]];
                    }
                    XElement ShapeXML = new XElement("Shape", new XAttribute("Master", master), new XAttribute("Type", "Foreign"), new XAttribute("Name", name), new XAttribute("NameU", nameu), new XAttribute("ID", $"{index_reordered[j] + 1}"),
                                        new XElement("Cell", new XAttribute("V", pinX[j]), new XAttribute("N", "PinX")),
                                        new XElement("Cell", new XAttribute("V", pinY[j]), new XAttribute("N", "PinY")),
                                        new XElement("Cell", new XAttribute("V", $"{width}"), new XAttribute("N", "Width")),
                                        new XElement("Cell", new XAttribute("V", $"{height}"), new XAttribute("N", "Height")),
                                        new XElement("Section", new XAttribute("N", "Character"),
                                            new XElement("Row", new XAttribute("IX", "0"),
                                                new XElement("Cell", new XAttribute("V", $"{sf * textsize}"), new XAttribute("N", "Size")))),
                                        new XElement("Text", stenciltext));
                    pageXML.Root.Elements().ElementAt(0).Add(ShapeXML);
                }
                //Add a Connects Element
                string text = "";
                List <string> texts = new List<string>();
                int connectind = 1;
                bool alreadyconnected = false;
                bool connectedwithdifferentinterface = false;
                XElement ConnectXML = new XElement("Connect");
                XElement ConnectsXML = new XElement("Connects");
                List<XElement> ConnectShapesXML = new List<XElement>();
                
                for (int j = 0; j < index_reordered.Count(); j++) {
                    if (interfaces[index_reordered[j]] == null) {
                        continue;
                    }

                    for (int m = 0; m < interfaces[index_reordered[j]].Length / 3; m++) {
                        text = $"{interfaces[index_reordered[j]][m, 0]} - {interfaces[index_reordered[j]][m, 1]}"; // local port - neighbor port (outgoing)
                        XElement ConnectionStartExists = Backend.GetXElementByAttribute(ConnectsXML.Elements(), "ToSheet", $"{index_reordered[j] + 1}");
                        XElement ConnectionEndExists = new XElement("cell");
                        if (ConnectionStartExists != null) {
                            // So there is already a connection for this device, but we don't know where to yet. Check all connections: starting device
                            IEnumerable<XElement> ToSheetElements = Backend.GetXElementsByAttribute(ConnectsXML.Elements(), "ToSheet", $"{index_reordered[j] + 1}");
                            for (int n = 0; n < ToSheetElements.Count(); n++) {
                                IEnumerable<XElement> FromSheetElements = Backend.GetXElementsByAttribute(ConnectsXML.Elements(), "FromSheet", ToSheetElements.ElementAt(n).Attribute("FromSheet").Value);
                                ConnectionEndExists = Backend.GetXElementByAttribute(FromSheetElements, "ToSheet", $"{hostname.FindIndex(x => x.Contains(interfaces[index_reordered[j]][m, 2])) + 1}");
                                // but are they connecting to each other?
                                if (ConnectionEndExists != null) {
                                    connectedwithdifferentinterface = true;
                                    // last check: are the interface names different
                                    // Check where 'texts' contains from [0] and to [1] then see interface description
                                    if (texts.Contains($"{index_reordered[j] + 1} " + $"{hostname.FindIndex(x => x.Contains(interfaces[index_reordered[j]][m, 2])) + 1} " + text)) {
                                        alreadyconnected = true;
                                    }
                                    break;
                                }
                            }
                            if (alreadyconnected) {
                                // the connection is there, no need to create another one
                                alreadyconnected = false;
                                connectedwithdifferentinterface = false;
                                continue;
                            }
                            if (connectedwithdifferentinterface) {
                                // First get ID of connector for updating text from ConnectsXML
                                foreach (XElement connector in ConnectShapesXML) {
                                    if (connector.Attribute("ID").Value == ConnectionEndExists.Attribute("FromSheet").Value) {
                                        connector.Element("Text").Value = connector.Element("Text").Value + "\n" + text;
                                        connectedwithdifferentinterface = false;
                                        continue;
                                    }
                                }
                                texts.Add($"{index_reordered[j] + 1} " + $"{hostname.FindIndex(x => x.Contains(interfaces[index_reordered[j]][m, 2])) + 1} " + text);
                                texts.Add($"{hostname.FindIndex(x => x.Contains(interfaces[index_reordered[j]][m, 2])) + 1} " + $"{index_reordered[j] + 1} " + $"{interfaces[index_reordered[j]][m, 1]} - {interfaces[index_reordered[j]][m, 0]}");
                                continue;
                            }
                        }
                        texts.Add($"{index_reordered[j] + 1} " + $"{hostname.FindIndex(x => x.Contains(interfaces[index_reordered[j]][m, 2])) + 1} " + text);
                        texts.Add($"{hostname.FindIndex(x => x.Contains(interfaces[index_reordered[j]][m, 2])) + 1} " + $"{index_reordered[j] + 1} " + $"{interfaces[index_reordered[j]][m, 1]} - {interfaces[index_reordered[j]][m, 0]}");
                        // 1. Add the Shape object
                        XElement ConnectShapeXML = new XElement("Shape", new XAttribute("Master", "1024"), new XAttribute("Type", "Shape"), new XAttribute("Name", "Dynamic connector"), new XAttribute("NameU", "Dynamic connector"), new XAttribute("ID", $"{hostname.Count() + connectind}"),
                            new XElement("Text", text),
                            new XElement("Cell", new XAttribute("N", "ShapeRouteStyle"), new XAttribute("V", "16")),
                            new XElement("Cell", new XAttribute("N", "BegTrigger"), new XAttribute("V", "2"), new XAttribute("F", $"_XFTRIGGER(Sheet.{index_reordered[j] + 1}!EventXFMod)")),
                            new XElement("Cell", new XAttribute("N", "EndTrigger"), new XAttribute("V", "2"), new XAttribute("F", $"_XFTRIGGER(Sheet.{hostname.FindIndex(x => x.Contains(interfaces[index_reordered[j]][m, 2])) + 1}!EventXFMod)")),
                            new XElement("Section", new XAttribute("N", "Character"),
                                new XElement("Row", new XAttribute("IX", "0"),
                                new XElement("Cell", new XAttribute("V", $"{sf * textsize}"), new XAttribute("N", "Size"), new XAttribute("U", "PT"))))); ;
                            //new XElement("Cell", new XAttribute("N", "ConFixedCode"), new XAttribute("V", "5")));

                        ConnectShapesXML.Add(ConnectShapeXML);
                        ConnectXML = new XElement("Connect",
                            new XAttribute("FromSheet", $"{hostname.Count() + connectind}"),
                            new XAttribute("FromCell", "BeginX"),
                            new XAttribute("FromPart", "12"),
                            new XAttribute("ToSheet", $"{index_reordered[j] + 1}"),
                            new XAttribute("ToCell", "PinX"),
                            new XAttribute("ToPart", "3"));
                        ConnectsXML.Add(ConnectXML);
                        ConnectXML = new XElement("Connect",
                            new XAttribute("FromSheet", $"{hostname.Count() + connectind}"),
                            new XAttribute("FromCell", "EndX"),
                            new XAttribute("FromPart", "9"),
                            new XAttribute("ToSheet", $"{hostname.FindIndex(x => x.Contains(interfaces[index_reordered[j]][m, 2])) + 1}"),
                            new XAttribute("ToCell", "PinX"),
                            new XAttribute("ToPart", "3"));
                        ConnectsXML.Add(ConnectXML);
                        connectind++;
                    }
                }
                for (int j = 0; j < ConnectShapesXML.Count(); j++) { 
                    // This had to be in a separate for loop.. to look back at the connector description and append additional connections in one go
                    pageXML.Root.Elements().ElementAt(0).Add(ConnectShapesXML[j]);
                }
                pageXML.Root.Add(ConnectsXML);
                Backend.SaveXDocumentToPart(pagePart, pageXML);
            }
        }
    }
}