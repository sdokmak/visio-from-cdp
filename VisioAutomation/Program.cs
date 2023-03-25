// Toggle Debug mode here
//#define DEBUG

using System;
using System.Collections.Generic;
using System.IO.Packaging;

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

namespace VisioAutomation {
    class Program {
        private static void Main() {
            GlobalVariables.stopwatch.Start();
            (List<string> hostname, List<string> devicetype, List<string> ipaddress, 
                List<string> model, List<string> site, List<int> tier, List<int> indtier, string[][,] interfaces) = Parser.Prepare_Data();
            string templatePath = "C:\\Users\\saber\\visodlldemo\\NTT Visio Template - A3-Landscape.vstx";
            string outputPath = "C:\\Users\\saber\\visodlldemo\\Generated Devices.vsdx";
            Backend.ConvertTemplateToDoc(templatePath, outputPath);
            using (Package visioPackage = Backend.OpenPackage(outputPath)) {
                //Backend.IteratePackageParts(visioPackage);
                Frontend.EmptyPages(visioPackage, site);
                Frontend.Add_CDP_Devices(visioPackage, hostname, devicetype, ipaddress, model, site, tier, indtier, interfaces);
                VisioData.SummaryPage(visioPackage);
                Backend.RecalcDocument(visioPackage);
            }
            GlobalVariables.stopwatch.Stop();
            Console.WriteLine("");
            Console.WriteLine("Elapsed time: " + GlobalVariables.stopwatch.ElapsedMilliseconds / 1000f + " seconds");
            Console.WriteLine("Press any key...");
            Console.ReadKey();
        }
    }
}