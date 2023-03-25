using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace VisioAutomation {
    // Parses the text files and figures out how the devices connect
    internal class Parser {
        private static VisioData.CDPParseDevice ParseDevice() {
            List<(string, string)> ParseDeviceList = new List<(string, string)>(); // type, regex
            VisioData.CDPParseDevice ParseDevice = new VisioData.CDPParseDevice();
            ParseDeviceList.Add(("Switch", "^(C1)*(-)*(WS)*[-]*C([1-3]|9[2-4])|^HX|^DS|^IE|CAT"));
            ParseDeviceList.Add(("Core Switch", "^(C1)*(-)*(WS)*[-]*C([4-6]|9[5-7])|^WS[-]*X45|^NEXUS[0-9]+|^N[1-9]K|^DS|NEX"));
            ParseDeviceList.Add(("WLC", "^C1[-]AIR|^AIR-CT"));
            ParseDeviceList.Add(("Router", "^CISCO[0-9][0-9][0-9][0-9]|^C[0-8][0-9][0-9][0-9]|^ASR|^ISR|^VEDGE"));
            ParseDeviceList.Add(("Access Point", "^AIR-AP"));
            ParseDeviceList.Add(("Server", "^UCS|^B230|^N[1-2][0-9]|^Fabric Extender|^CPS|^SNS"));
            ParseDeviceList.Add(("Firewall", "^FPR|^FMC"));
            ParseDeviceList.Add(("Voice Gateway", "^VG"));
            ParseDeviceList.Add(("Access Server", "^CSAC"));
            ParseDeviceList.Add(("NAC Appliance", "^ASA"));
            ParseDeviceList.Add(("VM", "^(R-|L-|)ISE-VM"));
            ParseDevice.List = ParseDeviceList;
            return ParseDevice;
        }
        private static string[] GetTextandLogs() {
            List<string> txtfiles = new List<string>(Directory.GetFiles(GlobalVariables.dir, "*.txt")).ToList();
            List<string> logfiles = new List<string>(Directory.GetFiles(GlobalVariables.dir, "*.log")).ToList();
            txtfiles.AddRange(logfiles);
            var files = txtfiles.ToArray();
            return files;
        }
        private static string[] GetHostnames() {
            var files = Parser.GetTextandLogs();
            string[] hostnames = new string[files.Length];
            string regexHostnames = "[a-zA-Z0-9-*()+_%$#@!.^]*(?=[.])*(?=[.](txt|log))";
            for (int i = 0; i < files.Length; i++) {
                if (Regex.IsMatch(files[i], regexHostnames)) {
                    hostnames[i] = Regex.Match(files[i], regexHostnames).Value;
                }
                else {
                    Console.BackgroundColor = ConsoleColor.Red;
                    Console.Write(" Error:");
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.WriteLine("Does this file have a hostname? ->" + files[i]);
                    break;
                }
            }
            return hostnames;
        }
        private static VisioData.Devices GetDevices() {
            VisioData.CDPParseDevice parseDevice = ParseDevice();
            VisioData.Devices Devices = new VisioData.Devices();
            List<VisioData.Device> ListDevice = new List<VisioData.Device> { };
            List<(string, string, string)> interfacelink = new List<(string, string, string)> { };
            int ind = 0;
            Boolean port_recorded = false;
            string[] hostnames = GetHostnames();
            string[] files = GetTextandLogs();
            for (int i = 0; i < hostnames.Length; i++) {
                // Initialize devices
                VisioData.Device Device = new VisioData.Device();
                Device.hostname = hostnames[i];
                ListDevice.Add(Device);
            }
            int initial_hostnames_length = hostnames.Length;
            for (int i = 0; i < initial_hostnames_length; i++) {
                string logtext = File.ReadAllText(files[i]);
                if (Regex.IsMatch(logtext, "Device ID: ((?!(>|#))(.|\r|\n))*")) {
                    string CDPSection = Regex.Match(logtext, "Device ID: ((?!(>|#))(.|\r|\n))*").Value;
                    while (Regex.IsMatch(CDPSection, "Device ID: ((?!(>|#))(.|\r|\n))*")) {
                        string CDPSubsection = Regex.Match(CDPSection, "Device ID: ((?!(Device ID: |>|#))(.|\n|\r))*").Value;
                        /*
                            public string model;
                            public string manufacturer;
                            public string devicetype;
                            public string hostname;
                            public string ipaddress;
                            public string site;
                            public string description;
                            public string serial;
                            public Boolean stacked;
                         */
                        // need the hostname and model before if statement to determine if exists and devicetype respectively :D
                        string CDPhostname = Regex.Match(CDPSubsection, "(?<=Device ID: ).*").Value.Trim('\r');
                        // sometimes hostname includes the domain name, remove that... but keep if it's a mac address!
                        if ((Regex.IsMatch(CDPhostname, "[.]")) && (!Regex.IsMatch(CDPhostname, "[0-9A-Z][0-9A-Z][0-9A-Z][0-9A-Z][.][0-9A-Z][0-9A-Z][0-9A-Z][0-9A-Z][.][0-9A-Z][0-9A-Z][0-9A-Z][0-9A-Z]"))) {
                            CDPhostname = Regex.Match(CDPhostname, "[^.]*(?=[.])").Value;
                        }
                        //if (hostnames[i] == "viclab-l19-dnac-bn01") {
                        //    Console.WriteLine("debug");
                        //}
                        string CDPmodel = Regex.Match(CDPSubsection, "(?<=Platform: cisco ).*(?=,)").Value;
                        string CDPipaddress = Regex.Match(CDPSubsection, "(?<=IP address: )[0-9.]*").Value;
                        string CDPtype = "";
                        for (int j = 0; j < parseDevice.List.Count; j++) {
                            if (Regex.IsMatch(CDPmodel, parseDevice.List[j].Item2)) {
                                CDPtype = parseDevice.List[j].Item1;
                                break;
                            }
                        }

                        if (hostnames.Any(CDPhostname.Contains)) {
                            // If device exists, update interface attribute and the rest of the fields if they're empty
                            ind = Array.IndexOf(hostnames, CDPhostname);
                            port_recorded = false;
                            interfacelink = new List<(string, string, string)> { };
                            if (ListDevice[ind].interfaces != null) {
                                interfacelink = ListDevice[ind].interfaces;
                                // If port connection already exists, ignore adding :)
                                for (int j = 0; j < interfacelink.Count; j++) {
                                    if (interfacelink[j].Item1 == Regex.Match(CDPSubsection, @"(?<= Port ID \(outgoing port\): ).*").Value.Trim('\r')) {
                                        port_recorded = true;
                                        break;
                                    }
                                }
                            }
                            if (!port_recorded) {
                                interfacelink.Add((Regex.Match(CDPSubsection, @"(?<= Port ID \(outgoing port\): ).*").Value.Trim('\r'), Regex.Match(CDPSubsection, "(?<=Interface: ).*(?=,)").Value, hostnames[i]));
                            }
                            ListDevice[ind].interfaces = interfacelink;
                            if (ListDevice[ind].model == null) {
                                ListDevice[ind].model = CDPmodel;
                                ListDevice[ind].devicetype = CDPtype;
                                ListDevice[ind].ipaddress = CDPipaddress;
                            }
                        }
                        else {
                            // if device doesn't exist, add one to the list for hostnames and devices
                            Array.Resize(ref hostnames, hostnames.Length + 1);
                            hostnames[hostnames.Length - 1] = CDPhostname;
                            interfacelink = new List<(string, string, string)> { };
                            interfacelink.Add((Regex.Match(CDPSubsection, @"(?<= Port ID \(outgoing port\): ).*").Value.Trim('\r'), Regex.Match(CDPSubsection, "(?<=Interface: ).*(?=,)").Value, hostnames[i]));
                            VisioData.Device Device = new VisioData.Device {
                                interfaces = interfacelink,
                                hostname = CDPhostname,
                                model = CDPmodel,
                                ipaddress = CDPipaddress,
                                devicetype = CDPtype
                            };
                            // Device.devicetype = ...;
                            ListDevice.Add(Device);
                        }
                        // and finally, add the interface connection to the current device
                        port_recorded = false;
                        interfacelink = new List<(string, string, string)> { };
                        if (ListDevice[i].interfaces != null) {
                            interfacelink = ListDevice[i].interfaces;
                            // If port connection already exists, ignore adding :)
                            for (int j = 0; j < interfacelink.Count; j++) {
                                if (interfacelink[j].Item1 == Regex.Match(CDPSubsection, "(?<=Interface: ).*(?=,)").Value) {
                                    port_recorded = true;
                                    break;
                                }
                            }
                        }
                        if (!port_recorded) {
                            interfacelink.Add((Regex.Match(CDPSubsection, "(?<=Interface: ).*(?=,)").Value, Regex.Match(CDPSubsection, @"(?<= Port ID \(outgoing port\): ).*").Value.Trim('\r'), CDPhostname));
                        }
                        ListDevice[i].interfaces = interfacelink;
                        CDPSection = CDPSection.Substring(CDPSubsection.Length);
                    }
                }
                else {
                    Console.BackgroundColor = ConsoleColor.Yellow;
                    Console.ForegroundColor = ConsoleColor.Black;
                    Console.Write(" Warning:");
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine(" No 'show cdp neighbors detail' command in log: " + files[i]);
                    break;
                }
            }
            for (int i = 0; i < ListDevice.Count; i++) {
                if (ListDevice[i].interfaces != null) {
                    int[] interfacekeys = new int[ListDevice[i].interfaces.Count];
                    for (int j = 0; j < interfacekeys.Length; j++) {
                        string interfacekey = Regex.Replace(ListDevice[i].interfaces[j].Item1, @"[^\d]", "");
                        // to make 2/0/1 bigger than 1/0/41 for example.. 
                        if (interfacekey.Length < 4) {
                            interfacekey = interfacekey.Substring(0, interfacekey.Length - 1) + "0" + interfacekey.Substring(interfacekey.Length - 1);
                        }
                        interfacekeys[j] = Int16.Parse(interfacekey);
                    }
                    (string, string, string)[] intlinkarray = ListDevice[i].interfaces.ToArray();
                    Array.Sort(interfacekeys, intlinkarray);
                    ListDevice[i].interfaces = intlinkarray.ToList();
                }
            }
            for (int i = 0; i < ListDevice.Count; i++) {
                if (ListDevice[i].interfaces != null) {
                    string[] neighbors = new string[ListDevice[i].interfaces.Count];
                    for (int j = 0; j < neighbors.Length; j++) {
                        neighbors[j] = ListDevice[i].interfaces[j].Item3;
                    }
                    ListDevice[i].count_neighbors = neighbors.Distinct().Count();
                }
                else {
                    ListDevice[i].count_neighbors = 0;
                }
                Console.WriteLine(ListDevice[i].hostname + ": " + ListDevice[i].count_neighbors.ToString());
            }
            Devices.Device = ListDevice;
            return Devices;
        }
        private static VisioData.Devices AddDeviceTiers(VisioData.Devices Devices_no_tier) {
            VisioData.Devices Devices = Devices_no_tier;
            List<VisioData.Device> ListDevice = Devices.Device;
            string[] hostnames = new string[ListDevice.Count];
            int[] count_neighbors = new int[ListDevice.Count];
            string[] devicetype = new string[ListDevice.Count];
            for (int i = 0; i < ListDevice.Count; i++) {
                hostnames[i] = ListDevice[i].hostname;
                count_neighbors[i] = ListDevice[i].count_neighbors;
                devicetype[i] = ListDevice[i].devicetype;
            }
            // [ ] each group of connections is a tab
            // [ ] start with top tier being 1 and iterate until you reach the APs
            // [ ] routers are always tier 1
            // [ ] anything connected to routers are tier 2
            // [ ] Find a router as starting point
            // [ ] If no router, find a core switch
            // [ ] If no router - free form, pick the one with most connections first
            // if two pick two
            //string[] irouters = new string[0];
            string[] icoreswitches = new string[0];
            int[] iarray = new int[0];
            int[][] ijkarray = new int[1][];
            string[] neighbors = new string[0];
            int[] jarray = new int[0];
            // Router as starting node if there is one.
            string[] remaininghostnames = new string[0];
            string[] remainingdevicetypes = new string[0];
            int[] remainingcountneighbors = new int[0];
            int[] remainingindices = new int[0];
            int site = 1;
            int loop = 0;
            int[] ijkrow = new int[0];
            int endloop = 0;
            int[] temparray = new int[0];
            int jstart = 0;
            int istart = 0;
            ijkarray[0] = new int[] { Array.IndexOf(devicetype, "Router") };
            //ijkarray[0] = devicetype.Select((value, index) => new { value, index })
            //    .Where(x => x.value == "Router")
            //    .Select(x => x.index)
            //    .ToArray();
            if (ijkarray[0][0] == -1) {
                ijkarray[0] = new int[] { Array.IndexOf(devicetype, "Core Switch") };
                //// Otherwise, Core switch as starting node if there is one
                //ijkarray[0] = devicetype.Select((value, index) => new { value, index })
                //    .Where(x => x.value == "Core Switch")
                //    .Select(x => x.index)
                //    .ToArray();
                if (ijkarray[0][0] == -1) {
                    ijkarray[0] = new int[] { Array.IndexOf(count_neighbors, count_neighbors.Max()) };
                    //// Otherwise, most neighbors as starting node
                    //ijkarray[0] = count_neighbors.Select((value, index) => new { value, index })
                    //    .Where(x => x.value == count_neighbors.Max())
                    //    .Select(x => x.index)
                    //    .ToArray();
                }
            }
            string[] allneighbors = new string[0];
            while (true) {
                int tier = 1;
                if (loop != 0) {
                    remaininghostnames = hostnames.Except(allneighbors).ToArray();
                    if (remaininghostnames.Length == 0) {
                        break;
                    }
                    Array.Resize(ref remainingindices, remaininghostnames.Length);
                    Array.Resize(ref remainingdevicetypes, remaininghostnames.Length);
                    Array.Resize(ref remainingcountneighbors, remaininghostnames.Length);
                    for (int i = 0; i < remaininghostnames.Length; i++) {
                        remainingindices[i] = Array.IndexOf(hostnames, remaininghostnames[i]);
                        remainingdevicetypes[i] = devicetype[remainingindices[i]];
                        remainingcountneighbors[i] = count_neighbors[remainingindices[i]];
                    }
                    // This is a repeat of the initialization
                    temparray = new int[] { Array.IndexOf(remainingdevicetypes, "Router") };
                    if (temparray[0] == -1) {
                        temparray = new int[] { Array.IndexOf(remainingdevicetypes, "Core Switch") };
                        if (temparray[0] == -1) {
                            temparray = new int[] { Array.IndexOf(remainingcountneighbors, remainingcountneighbors.Max()) };
                            if (remainingindices.Length == 1) {
                                temparray = new int[] { remainingindices[0] };
                                if (ListDevice[temparray[0]].tier != 0) {
                                    endloop = 1;
                                }
                            }
                        }
                    }
                    if (endloop == 1) {
                        break;
                    }
                    istart = ijkarray[tier - 1].Length;
                    Array.Resize(ref ijkarray[tier - 1], ijkarray[tier - 1].Length + 1);
                    ijkarray[tier - 1][istart] = temparray[0];
                }
                while (tier < ijkarray.Length + 1) {
                    string[] iallneighbors = new string[0];
                    ijkrow = new int[0];
                    for (int i = istart; i < ijkarray[tier - 1].Length; i++) {
                        //if (ListDevice[ijkarray[tier-1][i]].tier == 0 || ListDevice[ijkarray[tier - 1][i]].tier == tier) {
                        ListDevice[ijkarray[tier - 1][i]].tier = tier;
                        ListDevice[ijkarray[tier - 1][i]].site = site.ToString();
                        // Site Neighbors: This won
                        //}
                        //else {
                        //    ijkarray[tier-1] = ijkarray[tier-1].Where(x => x != ijkarray[tier-1][i]).ToArray();
                        //    if (i != 0) {
                        //        i--;
                        //    }
                        //}
                        if (ListDevice[ijkarray[tier - 1][i]].interfaces != null) {
                            neighbors = new string[ListDevice[ijkarray[tier - 1][i]].interfaces.Count];
                            for (int j = 0; j < neighbors.Length; j++) {
                                if (ListDevice[ijkarray[tier - 1][i]].tier == 0 || ListDevice[ijkarray[tier - 1][i]].tier == tier) {
                                    neighbors[j] = ListDevice[ijkarray[tier - 1][i]].interfaces[j].Item3;
                                }
                            }
                            // 1. neighbors: neighbors for one device in "i" for loop
                            // 2. allneighbors: everything that's been or will be recorded already in ijkarray
                            // 3. iallneighbors: all neighbors after passing through all connected devices for that 'site'
                            // 4. ijkrow: a single row in ijk array, should range from previous size to new size
                            Array.Resize(ref allneighbors, allneighbors.Length + neighbors.Length);
                            neighbors.CopyTo(allneighbors, allneighbors.Length - neighbors.Length);
                            Array.Resize(ref iallneighbors, iallneighbors.Length + neighbors.Length);
                            neighbors.CopyTo(iallneighbors, iallneighbors.Length - neighbors.Length);
                        }
                        iallneighbors = iallneighbors.Distinct().ToArray();
                        // Get indices to add into ijkarray
                        Array.Resize(ref ijkrow, iallneighbors.Length);
                        for (int j = 0; j < iallneighbors.Length; j++) {
                            ijkrow[j] = Array.IndexOf(hostnames, iallneighbors[j]);
                        }
                    }
                    //if (ijkrow.Length == 0) {
                    //    continue;
                    //}
                    //if (tier < ijkarray.Length) { 
                    //    // get indices in hostname array

                    //}
                    if (ijkarray.Length > tier) {
                        jstart = ijkarray[tier].Length;
                    }
                    else {
                        jstart = 0;
                    }
                    // exclude devices from previous tiers
                    for (int j = 0; j < tier; j++) {
                        ijkrow = ijkrow.Except(ijkarray[tier - 1 - j]).ToArray();
                    }
                    // clean up repeating
                    allneighbors = allneighbors.Distinct().ToArray();
                    if (ijkrow.Length == 0) {
                        break;
                    }
                    if (loop == 0 || ijkarray.Length < tier) {
                        Array.Resize(ref ijkarray, ijkarray.Length + 1);
                        ijkarray[tier] = ijkrow;
                    }
                    else {
                        istart = ijkarray[tier].Length;
                        Array.Resize(ref ijkarray[tier], ijkrow.Length + istart);
                        ijkrow.CopyTo(ijkarray[tier], istart);
                    }
                    tier++;
                }
                site++;
                loop++;
            }
            List<int> newijk = new List<int> { };
            List<int> usedneighbors = new List<int> { };
            int[] indneighbors = new int[0];
            List<int> tierindneighbors = new List<int> { };
            int indneighbor = -1;
            // now we know how many, For loops are back!!
            for (int i = 0; i < ijkarray.Length; i++) {
                // Get device neighbors for the same tier
                newijk = ijkarray[i].ToList();
                for (int j = 0; j < ijkarray[i].Length - 1; j++) {
                    indneighbors = new int[count_neighbors[ijkarray[i][j]]];
                    for (int k = 0; k < indneighbors.Length; k++) {
                        indneighbors[k] = Array.IndexOf(hostnames, ListDevice[ijkarray[i][j]].interfaces[k].Item3);
                    }
                    indneighbors = indneighbors.Distinct().ToArray();
                    // if neighbor in tier
                    tierindneighbors = new List<int> { };
                    for (int k = 0; k < indneighbors.Length; k++) {
                        if (ijkarray[i].Contains(indneighbors[k])) {
                            tierindneighbors.Add(indneighbors[k]);
                        }
                    }
                    if (tierindneighbors.Count != 0 && !tierindneighbors.Contains(ijkarray[i][j + 1])) {
                        if (usedneighbors.Count != 0) {
                            indneighbor = -1;
                            for (int k = 0; k < tierindneighbors.Count; k++) {
                                if (!usedneighbors.Contains(tierindneighbors[k])) {
                                    indneighbor = tierindneighbors[k];
                                }
                            }
                            if (indneighbor == -1) {
                                continue;
                            }
                        }
                        else {
                            indneighbor = tierindneighbors[0];
                        }
                        usedneighbors.Add(indneighbor);
                        Backend.Swap(newijk, Array.IndexOf(ijkarray[i], indneighbor), j + 1);
                    }
                    ijkarray[i] = newijk.ToArray();
                }
            }
            for (int i = 0; i < ijkarray.Length; i++) {
                for (int j = 1; j <= ijkarray[i].Length; j++) {
                    ListDevice[ijkarray[i][j - 1]].tierind = j;
                }
            }


            Devices.Device = ListDevice;
            return Devices;
        }
        public static (List<string>, List<string>, List<string>, List<string>, List<string>, List<int>, List<int>, string[][,]) Prepare_Data() {
            VisioData.Devices Devices = AddDeviceTiers(GetDevices());
            List<VisioData.Device> ListDevice = Devices.Device;
            List<string> hostname = new List<string> { };
            List<string> devicetype = new List<string> { };
            List<string> ipaddress = new List<string> { };
            List<string> model = new List<string> { };
            List<string> site = new List<string> { };
            List<int> tier = new List<int> { };
            List<int> tierind = new List<int> { };
            string[][,] interfaces = new string[ListDevice.Count][,];


            for (int i = 0; i < ListDevice.Count; i++) {
                if (ListDevice[i].hostname != null) {
                    hostname.Add(ListDevice[i].hostname);
                }
                else {
                    hostname.Add("No hostname");
                }
                if (ListDevice[i].devicetype != null) {
                    devicetype.Add(ListDevice[i].devicetype);
                }
                else {
                    devicetype.Add("Question Mark");
                }
                if (ListDevice[i].ipaddress != null) {
                    ipaddress.Add(ListDevice[i].ipaddress);
                }
                else {
                    ipaddress.Add("No IP");
                }
                if (ListDevice[i].model != null) {
                    model.Add(ListDevice[i].model);
                }
                else {
                    model.Add("No model");
                }
                if (ListDevice[i].site != null) {
                    site.Add(ListDevice[i].site);
                }
                else {
                    site.Add("No site");
                }
                tier.Add(ListDevice[i].tier);
                tierind.Add(ListDevice[i].tierind);
                if (ListDevice[i].interfaces != null) {
                    // checkpoint, will need to add the if statement for RemoveRepeatingConnections here, cuz the array is hard to breakdown after it's defined.
                    interfaces[i] = new string[ListDevice[i].interfaces.Count, 3];
                    for (int j = 0; j < interfaces[i].Length / 3; j++) {
                        interfaces[i][j, 0] = ListDevice[i].interfaces[j].Item1;
                        interfaces[i][j, 1] = ListDevice[i].interfaces[j].Item2;
                        interfaces[i][j, 2] = ListDevice[i].interfaces[j].Item3;
                    }
                }
            }
            // update devicecount array to show summary page
            for (int i = 0; i < devicetype.Count; i++) {
                VisioData.deviceCount[Array.IndexOf(VisioData.DeviceTypes, devicetype[i])]++;
            }
            interfaces = VisioData.ShortenInterfaceNaming(interfaces);
            return (hostname, devicetype, ipaddress, model, site, tier, tierind, interfaces);
        }
    }
}