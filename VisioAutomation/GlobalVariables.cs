// Toggle Debug mode here
// #define DEBUG

using System.IO;
using System;
using System.Diagnostics;

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
    class GlobalVariables {
        //public static readonly string templateFilename = "Visio Template - A3-Landscape.vstx";
        //public static readonly string generatedFilename = "Generated Devices.vsdx";
        //public static readonly string dir = Environment.CurrentDirectory; // for .exe builds, uncomment and comment the below line
        public static readonly string dir = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName;
        public static readonly Stopwatch stopwatch = new Stopwatch();
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