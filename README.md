# Visio from CDP
This script gives you a clean visio diagram for any cisco devices that run 'show cdp nei'. Simply run the script with the 'show cdp nei' logs and it'll do the rest.

## Connection automation
The idea is to fill up the connections to the most accurate degree possible based on the given info.
- The executable is able to create connections of neighbors that don't have cdp logs of their own, if they exist in other logs.
- Using a top-down architecture, the first row would be for any routers. No routers > Core switch. No Core switch > Device with most connections.
- After first row, it looks at the neighbors for each first row device and populates their neighbors in a for loop.
- If there are any 
- Every group of connections are a separate site.


<!-- ## To-DO (old)
### Archive
- [x] Update Add_CDP_Devices method
- [x] update ijkarray to position neighboring devices to one another
    - [x] neighboring tiers
    - [x] same tier
- [x] Add "tier" attribute to devices
    - [x] routers are always tier 1
    - [x] anything connected to routers are tier 2 and so on
- [x] record interface connection
    - [x] Go to neighbor hostname and check if outgoing port exists in ingoing port list.
- [x] add all hostnames as new devices
- [x] if filled in already, only append interface connection list.
- [x] Grab hostname from textfile name
- [x] Set up structure and method of approach
    - should be simple and intuitive
    - click of a button
    - utilizing both cdp nei outputs and excel sheet
        - excel sheet is for the details of each and every device you will be running cdp nei on
        - 'cdp nei' outputs are textfiles, either .log or .txt. tool parses text file to get all devices and connects the dots.
        - the console application should be in the same folder as all this data... Intuitive!!!
    - [x] have a look at previous cdp logs: SVHA, woolies, L19


### Stategies for node-link alorithm
#### Senario 1: Router(s) detected, attempt as top tier
- Phase 1:
    1. Create array of all routers to sweep through and check neighboring connections
        - Each router is assumed separate site 1 until proven otherwise.
- Phase 2:
    1. Record neighbors, assume tier 2 (unless connected to more than 2 tier 2 devices in later Phase).
- Phase 3:
    1. Check that not more than 2 neighbors are connected to more than 1 neighbor.
    2. If True - nothing more to do
    3. Else - for same tier connection, check which has more previous tier connections.
        - If equal, choose one alphanumerically. If one more than the other prioritise other one
        - utilize same tier spacing, but different tier.
- Phase 4:




## Considerations for future addons
- [ ] An option to include APs as well? "sho ap cdp neighbors all"
- [ ] different colors for different interface types? Fiber or ethernet.
- [ ] CDP doesn't recognized whether a switch is stacked - but this can be added if cdp detects connections of 2/x/x or higher
- [ ] Add visio template into script? -->

