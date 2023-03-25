// Toggle Debug mode here
//#define DEBUG

using System;
using System.Linq;

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
    class PrintToConsole {
        public static (int, int) DeviceFound(string manufacturer, string model, string devicetype, int i, int j, int endrange, int PrimaryAssets, int DevicesFound, int[] DeviceCount) {
            // Appends to device 
            Console.Write("{0:mm\\:ss} (", GlobalVariables.stopwatch.Elapsed);
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.Write((i-1).ToString());
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine("/"+ (endrange-1).ToString() + ") " + manufacturer + ", " + model + ", " + devicetype);
            DeviceCount[j]++;
            PrimaryAssets++;
            DevicesFound++;
            return (PrimaryAssets, DevicesFound);
        }
        public static int LicenseFound(string manufacturer, string model, string devicetype, int i, int j, int endrange, int PrimaryAssets, int[] DeviceCount) {
            // Appends to device 
            Console.Write("{0:mm\\:ss} (", GlobalVariables.stopwatch.Elapsed);
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.Write((i-1).ToString());
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine("/"+ (endrange-1).ToString() + ") " + manufacturer + ", " + model + ", " + devicetype);
            DeviceCount[j]++;
            PrimaryAssets++;
            return PrimaryAssets;
        }
        public static (int, int) ManufacturerIgnored(string manufacturer, string model, int i, int endrange, int PrimaryAssets, int ManufacturersNotRecognised) {
            Console.Write("{0:mm\\:ss} (", GlobalVariables.stopwatch.Elapsed);
            Console.BackgroundColor = ConsoleColor.DarkYellow;
            Console.Write((i-1).ToString());
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write("/"+ (endrange-1).ToString() + ") " + manufacturer + " (Not recognised)" + ", " + model);
            if (model == null) {
                Console.WriteLine("Null");
            }
            else {
                Console.WriteLine(model);
            }
            PrimaryAssets++;
            ManufacturersNotRecognised++;
            return (PrimaryAssets, ManufacturersNotRecognised);
        }
        public static int ComponentAsset(int i, int endrange, int ComponentAssets) {
            Console.WriteLine("{0:mm\\:ss} ("+ (i-1).ToString() + "/"+ (endrange-1).ToString() + ") Component asset", GlobalVariables.stopwatch.Elapsed);
            ComponentAssets++;
            return ComponentAssets;
        }
        public static (int, int) DeviceIgnored(string manufacturer, string model, int i, int endrange, int PrimaryAssets, int DevicesIgnored) {
            Console.Write("{0:mm\\:ss} (", GlobalVariables.stopwatch.Elapsed);
            Console.BackgroundColor = ConsoleColor.DarkGray;
            Console.Write((i-1).ToString());
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine("/"+ (endrange-1).ToString() + ") " + manufacturer + ", " + model + ", Not a device");
            PrimaryAssets++;
            DevicesIgnored++;
            return (PrimaryAssets, DevicesIgnored);
        }
        public static (int, int, string[]) DeviceNotRecognised(string manufacturer, string model, string description, int i, int endrange, int PrimaryAssets, int DevicesNotRecognised, string[] ArrayDevicesNotRecognised) {
            Console.Write("{0:mm\\:ss} (", GlobalVariables.stopwatch.Elapsed);
            Console.BackgroundColor = ConsoleColor.Red;
            Console.Write((i-1).ToString());
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine("/"+ (endrange-1).ToString() + ") " + manufacturer + ", " + model + ", " + description);
            PrimaryAssets++;
            DevicesNotRecognised++;
            ArrayDevicesNotRecognised = ArrayDevicesNotRecognised.Concat(new string[] { (i-1).ToString() }).ToArray();
            return (PrimaryAssets, DevicesNotRecognised, ArrayDevicesNotRecognised);
        }
        public static (int, int) ModelNameEmpty(string manufacturer, int i, int endrange, int PrimaryAssets, int ModelNameEmpty) {
            Console.Write("{0:mm\\:ss} (", GlobalVariables.stopwatch.Elapsed);
            Console.BackgroundColor = ConsoleColor.DarkMagenta;
            Console.Write((i-1).ToString());
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine("/"+ (endrange-1).ToString() + ") " + manufacturer + ", Null");
            PrimaryAssets++;
            ModelNameEmpty++;
            return (PrimaryAssets, ModelNameEmpty);
        }
        public static (int, int) NoIP(string manufacturer, string model, int i, int endrange, int PrimaryAssets, int NoIP) {
            Console.WriteLine("{0:mm\\:ss} (" + (i - 1).ToString() + "/" + (endrange - 1).ToString() + ") Primary asset with no IP, " + manufacturer + ", " + model, GlobalVariables.stopwatch.Elapsed);
            PrimaryAssets++;
            NoIP++;
            return (PrimaryAssets, NoIP);
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