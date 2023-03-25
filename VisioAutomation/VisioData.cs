// Toggle Debug mode here
//#define DEBUG

using System.Xml.Linq;
using System.IO;
using System;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text.Json;

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
    class VisioData {
        public static string[] uniquesites = { };
        public static readonly string[] DeviceTypes = new string[16] { "Core Switch", "Switch", "Router", "WLC", "Access Point", "Server", "Storage", "Firewall", "Voice Gateway", "Access Server", "NAC Appliance",
                                                                    "VM", "Network TAP", "ADC / Load Balancer", "Licenses and Subscriptions", "Question Mark"};
        public static readonly int[] MasterNumber = new int[16] { 1034, 1035, 1039, 1042, 1036, 1048, 1048, 1049, 1050, 1051, 1052, 1053, 1054, 1055, 1060, 1059 };
        public static readonly int[] DeviceWidth = new int[16] { 18, 19, 19, 19, 19, 18, 18, 10, 12, 19, 21, 14, 13, 14, 13, 6 };
        public static readonly int[] DeviceHeight = new int[16] { 26, 8, 11, 8, 8, 18, 18, 16, 12, 19, 15, 14, 13, 14, 12, 11 };
        public static int[] deviceCount = Backend.Initarray(DeviceTypes, 0);

        //public string RegexValue;
        //public string Manufacturer_Index;
        //public Access_Switch;
        //public Router;
        //public WLC;
        //public Server;
        //public Firewall;
        //public Voice_Gateway;
        //public Access_Control_Server;
        //public NAC_Appliance;
        //public VM;
        //public Network_TAP;
        //public ADC;
        //public License_and_Subscription;
        public static int PrimaryAssets = 0;
        public static int ComponentAssets = 0;
        public static int DevicesFound = 0;
        public static int ManufacturersNotRecognised = 0;
        public static int DeviceNotRecognised = 0;
        public static int ModelNameEmpty = 0;
        public class FailedAttempt {
            public int Timeout { get; set; }
            public string Status { get; set; }
        }
        public class IcmpInfo {
            public string Status { get; set; }
            public int Ttl { get; set; }
            public List<FailedAttempt> FailedAttempts { get; set; }
            public int RoundtripTime { get; set; }
        }
        public class Child {
            public string Id { get; set; }
            public string Value { get; set; }
        }
        public class Snmp {
            public string Id { get; set; }
            public List<Child> Children { get; set; }
            public string Name { get; set; }
            public string Value { get; set; }
        }
        public class DiscoveryResult {
            public string IPAddress { get; set; }
            public IcmpInfo IcmpInfo { get; set; }
            public List<Snmp> Snmp { get; set; }
        }
        public class FailedSnmpQueries {
        }
        public class Summary {
            public int IcmpHostsCount { get; set; }
            public int SnmpHostsSuccessCount { get; set; }
            public int SnmpHostsFailuresCount { get; set; }
            public FailedSnmpQueries FailedSnmpQueries { get; set; }
        }
        public class Json_Report {
            public string CustomerName { get; set; }
            public string ScanId { get; set; }
            public string Version { get; set; }
            public DateTime Started { get; set; }
            public DateTime Finished { get; set; }
            public string Elapsed { get; set; }
            public Summary Summary { get; set; }
            public List<DiscoveryResult> DiscoveryResults { get; set; }
        }
        public class Devices {
            public List<Device> Device { get; set; }
        }
        public class Device {
            public string model;
            public string devicetype;
            public string hostname;
            public string ipaddress;
            public string site;
            public int count_neighbors;
            public int tier;
            public int tierind;
            public List<(string, string, string)> interfaces { get; set; } // ingoing, outgoing, neighbor
        }
        public static string[][,] ShortenInterfaceNaming(string[][,] interfaces) {
            string[] intnamereg = {
                "^Bundle-Ether[^0-9]*",
                "^Bundle-POS[^0-9]*",
                "^FastEthernet[^0-9]*",
                "^FiftyGig[^0-9]*",
                "^FortyGig[^0-9]*",
                "^FourHundredGig[^0-9]*",
                "^GCC0[^0-9]*",
                "^GCC1[^0-9]*",
                "^GigabitEthernet[^0-9]*",
                "^HundredGig[^0-9]*",
                "^IMA[^0-9]*",
                "^InterflexLeft[^0-9]*",
                "^InterflexRight[^0-9]*",
                "^Loopback[^0-9]*",
                "^MgmtEth[^0-9]*",
                "^Multilink[^0-9]*",
                "^Null[^0-9]*",
                "^POS[^0-9]*",
                "^PW-Ether[^0-9]*",
                "^PW-IW[^0-9]*",
                "^SRP[^0-9]*",
                "^Serial[^0-9]*",
                "^TenGig[^0-9]*",
                "^TwentyFiveGig[^0-9]*",
                "^TwoHundredGig[^0-9]*",
                "^nVFabric-Gig[^0-9]*",
                "^nVFabric-HundredGig[^0-9]*",
                "^nVFabric-TenGig[^0-9]*",
                "^tunnel-ipsec[^0-9]*"
            };
            string[] intnameshort = {
                "BE",
                "BP",
                "Fa",
                "Fi",
                "Fo",
                "F",
                "G0",
                "G1",
                "Gi",
                "Hu",
                "IMA",
                "IL",
                "IR",
                "Lo",
                "Mg",
                "Ml",
                "Nu",
                "POS",
                "PE",
                "PI",
                "SRP",
                "Se",
                "Te",
                "TF",
                "TH",
                "nG",
                "nH",
                "nT",
                "tsec"
            };
            // shorten naming convention for interfaces

            for (int i = 0; i < interfaces.Count(); i++) {
                // 'from' and 'to' interface names are in index 0 and 1
                if (interfaces[i] != null) { 
                    for (int j = 0; j < interfaces[i].GetLength(0); j++) {
                        for (int k = 0; k < 2; k++) {
                            for (int l = 0; l < intnamereg.Length; l++) { 
                                if (Regex.IsMatch(interfaces[i][j,k], intnamereg[l])) {
                                    interfaces[i][j, k] = Regex.Replace(interfaces[i][j, k], intnamereg[l], intnameshort[l]);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return interfaces;
        }
        public class CDPParseDevice {
            public List<(string, string)> List { get; set; } // type, regex
        }
        public static XDocument PageRelXML() {
            XNamespace relfileNS = "http://schemas.openxmlformats.org/package/2006/relationships";
            XDocument relfileXML = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XElement(relfileNS + "Relationships"));
            return relfileXML;
        }
        public static XDocument EmptyPage() {
            XDocument doc = XDocument.Parse("<?xml version='1.0' encoding='utf-8' ?><PageContents xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xml:space='preserve'/>");
            return doc;
        }
        public static void SummaryPage(Package visioPackage) {
            PackagePart documentPart = Backend.GetPackagePart(visioPackage, "http://schemas.microsoft.com/visio/2010/relationships/document");
            PackagePart pagesPart = Backend.GetPackagePart(visioPackage, documentPart, "http://schemas.microsoft.com/visio/2010/relationships/pages");
            // Change "Page-1" to "Summary" in Pages.xml
            XDocument pagesXML = Backend.GetXMLFromPart(pagesPart);
            pagesXML.Root.Elements().ElementAt(1).Attribute("Name").Value = "Summary";
            pagesXML.Root.Elements().ElementAt(1).Attribute("NameU").Value = "Summary";
            Backend.SaveXDocumentToPart(pagesPart, pagesXML);
            // Create Empty Summary Page
            XDocument SummaryPageXML = VisioData.EmptyPage();
            SummaryPageXML.Root.Add(new XElement("Shapes"));
            //List<string> TextFields = new List<string> { };
            //for (int i = 0; i < uniquesites.Length; i++) {
            //    TextFields.Add(uniquesites[i]);
            //}
            //double PinX = 2.657480257555857;
            //double PinY = 10.9842520156;
            //for (int i=0; i < TextFields.Count; i++) {
            //    SummaryPageXML.Root.Elements().ElementAt(0).Add(FieldTextBox(TextFields[i], i+1, PinX, PinY-(i+1)*0.5511811047273199));
            //}

            // Following the Powershell Script for organising devices
            int limit=8;
            double sf=1;
            var index = deviceCount.Select((value, index) => new { value, index })
                    .Where(x => x.value != 0)
                    .Select(x => x.index)
                    .ToList();
            (string master, string name, string nameu, string id) = ("", "", "", "");
            (double width, double height, double textsize) = (0, 0, 0.01388*8);
            int rows = (int)Math.Floor(Convert.ToDouble((index.Count)/limit));
            int[] devicesInRow = Enumerable.Repeat(limit, rows+1).ToArray();
            int[] row_index = { };
            for (int r = 0; r <= rows; r++) {
                row_index = row_index.Concat(Enumerable.Repeat(r, limit)).ToArray();
            }
            devicesInRow[devicesInRow.Length - 1] = (index.Count)%devicesInRow.Last();
            double[] pinX = { };
            double[] pinY = { };
            int k = 0;
            for (int j = 0; j < index.Count; j++) {
                if (j/limit != k) {
                    k++;
                }
                pinX = pinX.Concat(Enumerable.Repeat(8.27 + (sf * ((-0.75*(devicesInRow[k] - 1)) + 1.5*j - 1.5*(limit * row_index[j]))), 1)).ToArray();
                pinY = pinY.Concat(Enumerable.Repeat(5.26 + (sf * (1.5*row_index[j]-0.75*(rows))), 1)).ToArray();
            }
            double previousPinY = pinY.Last();
            for (int j = 0; j < index.Count; j++) {
                (master, name, nameu) = (MasterNumber[index[j]].ToString(), DeviceTypes[index[j]], DeviceTypes[index[j]]);
                (width, height) = (sf*0.03936*DeviceWidth[index[j]], sf*0.03936*DeviceHeight[index[j]]);
                //                                                                                                                                                                                          j+TextFields.Count+1
                XElement ShapeXML = new XElement("Shape", new XAttribute("Master", master), new XAttribute("Type", "Foreign"), new XAttribute("Name", name), new XAttribute("NameU", nameu), new XAttribute("ID", $"{j+1}"),
                                    new XElement("Cell", new XAttribute("V", pinX[j]), new XAttribute("N", "PinX")),
                                    new XElement("Cell", new XAttribute("V", pinY[j]), new XAttribute("N", "PinY")),
                                    new XElement("Cell", new XAttribute("V", $"{width}"), new XAttribute("N", "Width")),
                                    new XElement("Cell", new XAttribute("V", $"{height}"), new XAttribute("N", "Height")),
                                    new XElement("Section", new XAttribute("N", "Character"),
                                        new XElement("Row", new XAttribute("IX", "0"),
                                            new XElement("Cell", new XAttribute("V", $"{sf*textsize}"), new XAttribute("N", "Size")))),
                                    new XElement("Text", DeviceTypes[index[j]].ToString() + ": " + deviceCount[index[j]].ToString()));
                SummaryPageXML.Root.Elements().ElementAt(0).Add(ShapeXML);
            }




            PackagePart pagePart = Backend.GetPackagePart(visioPackage, pagesPart, "http://schemas.microsoft.com/visio/2010/relationships/page");
            using (Stream partStream = pagePart.GetStream(FileMode.Create, FileAccess.ReadWrite)) {
                SummaryPageXML.Save(partStream);
            }
            //Backend.SaveXDocumentToPart(pagePart, VisioData.EmptyPage());
            //XDocument checkpageXML = Backend.GetXMLFromPart(pagePart);
            //// Add the Summary
            //// 1. 
            
            //XDocument pageXML = Backend.GetXMLFromPart(pagePart);
            //pageXML = SummaryPageXML;
            //PackagePart checkpagePart = Backend.GetPackagePart(visioPackage, pagesPart, "http://schemas.microsoft.com/visio/2010/relationships/page");
        }
        public static PackagePart GetPageByrId(Package filePackage, PackagePart sourcePart, int rid) {
            PackageRelationship packageRel = sourcePart.GetRelationship($"rId{rid}");
            PackagePart page = null;
            if (packageRel != null) {
                Uri partUri = PackUriHelper.ResolvePartUri(sourcePart.Uri, packageRel.TargetUri);
                page = filePackage.GetPart(partUri);
            }
            return page;
        }
        public static XElement FieldTextBox(string text, int id, double pinX, double pinY) {
            XElement TextBoxXML = new XElement("Shape", new XAttribute("TextStyle", "3"), new XAttribute("LineSyle", "3"), new XAttribute("Type", "Shape"), new XAttribute("ID", id.ToString()),
                                    new XElement("Cell", new XAttribute("V", pinX.ToString()), new XAttribute("N", "PinX")),
                                    new XElement("Cell", new XAttribute("V", pinY.ToString()), new XAttribute("N", "PinY")),
                                    new XElement("Cell", new XAttribute("V", "3.149606312727553"), new XAttribute("N", "Width")),
                                    new XElement("Cell", new XAttribute("V", "0.5511811047273199"), new XAttribute("N", "Height")),
                                    new XElement("Cell", new XAttribute("V", "1.574803156363776"), new XAttribute("N", "LocPinX"), new XAttribute("F", "Width*0.5")),
                                    new XElement("Cell", new XAttribute("V", "0.2755905523636599"), new XAttribute("N", "LocPinY"), new XAttribute("F", "Height*0.5")),
                                    new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "Angle")),
                                    new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "FlipX")),
                                    new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "FlipY")),
                                    new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "ResizeMode")),
                                    new XElement("Cell", new XAttribute("V", "2"), new XAttribute("N", "ShapeShdwShow")),
                                    new XElement("Cell", new XAttribute("V", "102"), new XAttribute("N", "QuickStyleLineColor")),
                                    new XElement("Cell", new XAttribute("V", "102"), new XAttribute("N", "QuickStyleFillColor")),
                                    new XElement("Cell", new XAttribute("V", "102"), new XAttribute("N", "QuickStyleShadowColor")),
                                    new XElement("Cell", new XAttribute("V", "102"), new XAttribute("N", "QuickStyleFontColor")),
                                    new XElement("Cell", new XAttribute("V", "1"), new XAttribute("N", "QuickStyleLineMatrix")),
                                    new XElement("Cell", new XAttribute("V", "1"), new XAttribute("N", "QuickStyleFillMatrix")),
                                    new XElement("Cell", new XAttribute("V", "1"), new XAttribute("N", "QuickStyleEffectsMatrix")),
                                    new XElement("Cell", new XAttribute("V", "1"), new XAttribute("N", "QuickStyleFontMatrix")),
                                    new XElement("Section", new XAttribute("N", "Paragraph"),
                                        new XElement("Row", new XAttribute("IX", "0"),
                                            new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "HorzAlign")))),
                                    new XElement("Section", new XAttribute("N", "Geometry"), new XAttribute("IX", "0"),
                                        new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "NoFill")),
                                        new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "NoLine")),
                                        new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "NoShow")),
                                        new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "NoSnap")),
                                        new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "NoQuickDrag")),
                                        new XElement("Row", new XAttribute("IX", "1"), new XAttribute("T", "RelMoveTo"),
                                            new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "X")),
                                            new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "Y"))),
                                        new XElement("Row", new XAttribute("IX", "2"), new XAttribute("T", "RelLineTo"),
                                            new XElement("Cell", new XAttribute("V", "1"), new XAttribute("N", "X")),
                                            new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "Y"))),
                                        new XElement("Row", new XAttribute("IX", "3"), new XAttribute("T", "RelLineTo"),
                                            new XElement("Cell", new XAttribute("V", "1"), new XAttribute("N", "X")),
                                            new XElement("Cell", new XAttribute("V", "1"), new XAttribute("N", "Y"))),
                                        new XElement("Row", new XAttribute("IX", "4"), new XAttribute("T", "RelLineTo"),
                                            new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "X")),
                                            new XElement("Cell", new XAttribute("V", "1"), new XAttribute("N", "Y"))),
                                        new XElement("Row", new XAttribute("IX", "5"), new XAttribute("T", "RelLineTo"),
                                            new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "X")),
                                            new XElement("Cell", new XAttribute("V", "0"), new XAttribute("N", "Y")))),
                                    new XElement("Text",
                                        new XElement("pp", new XAttribute("IX", "0")), text));
            return TextBoxXML;
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