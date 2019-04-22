bom2basket

This python script combines several BOM files such as those generated by Altium Designer,
merges quantity and generates an excel file with 3 sheets: one for Digi-Key, one for Newark
and the last for the others.

As input, it takes an excel file with production data: each row contains the BOM file name,
the quantity of PCBs produced, A customer reference.

It will ask for a stock file, containing at least 2 columns, 'DPN' and 'Quantity' (order not important).
The buying quantities are the total quantities less the items in stock.
If a stock file is not selected, press cancel. The calculation will proceed without substracting the stock.

When it's done, it will ask for a filename for saving the basket.

Requirements:  
The BOM files are those generated by Altium, each one for a single PCB.

The columns needed are:  
Dist:	Distributors (Digi-Key, Newark, others)  
DPN:	Distributor part number  
Quantity: this column is the quantity of components needed for 1 PCB only.   

The order of the columns are not important, there may be different distributors.
Be sure to spell the distributor names correctly: "Digi-Key" and "Newark" specially.
Others aren't checked.


