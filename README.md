# Obmiar - AutoCAD addon helping with getting ducts and pipes quantities from drawings

### What is this?
The idea behind this project is to aid the process of measuring quantities of ventilation ducts and all kinds of pipes. 

### What it does?
The process is all about drawing polylines along the measured installations. Polylines are assigned to automatically created layers which names contain needed info (system name, material, diameter, insulation details etc.)
Drawn polylines are then exported to .sdf format which can be opened in MS Excel. Then you can join the data from multiple drawings and get Pivot Table report of all the measured installations.

### What is the difference
The main advantage of this method is the ease of use. You can create properly named layers with almost no effort. The script calculates lengths so it eliminates possibility of human error. It also helps in case of documentation revisions - you change the revised part of the drawing only and export the data again.

### How to use
* download ObmiarPW.dll located in Obmiar/bin/Debug
* open AutoCAD and use `NETLOAD` command to load the .dll file
* use `APPLOAD` command to load dependencies (listed in section below)
* now you can use `OBMIAR` command to show the add-on palette

![alt text](https://pawelwnuk.pl/images/BoQ1.png "Obmiar Add-on Interface")

### TO-DO and ideas
* color grade lines based on diameter of pipe
* draw measured polylines with corresponding system name (line style)
* multilanguage version
* better screen space use
* ...

### Caveats 
* Pay attention to drawing scale, in case of differences in scale it's best to rescale all the drawings to corresponding scale.
* If you use object snap while measuring be careful with the Z axis (especially on ventilations drawings). AutoCAD can snap one end of polyline on different Z-coordinate than the other end. It looks like there is no problem from the top view but results in wrong length calculation.

### Dependencies
LOS.lsp [https://www.theswamp.org/index.php?topic=53844.0](https://www.theswamp.org/index.php?topic=53844.0)
TLEN.lsp [http://www.lee-mac.com/totallengthandarea.html](http://www.lee-mac.com/totallengthandarea.html)
ADDLEN.vlx [https://www.cadforum.cz/cadforum_en/download.asp?fileID=1013](https://www.cadforum.cz/cadforum_en/download.asp?fileID=1013)

### Used and modified code
Creating a docking palette: [https://forums.autodesk.com/autodesk/attachments/autodesk/152/26712/1/CP205-2_Mike_Tuersley.pdf](https://forums.autodesk.com/autodesk/attachments/autodesk/152/26712/1/CP205-2_Mike_Tuersley.pdf)
Layer creation and assignment: [https://through-the-interface.typepad.com](https://through-the-interface.typepad.com)