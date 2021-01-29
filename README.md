# AutoHotKey
A set of tools written in the automation language AutoHotKey

Download the exe file and store it anywhere you like.

A set of tools can be started by pressing the F10 function key. A menu will appear.
Most tools are to be used in Notepad++ or TM1 Architect/Perspectives (notably, Turbo Integrator and the Rules Editor).

The corresponding AHK file contains the source code. Please keep my header in it.

I also put out my custom "user defined languages" in Notepad++. I add TM1 and AHK language syntax support.
The code in the exe file will be toggle it on automatically with some of the tools.

As Notepad++ is used and automated, we need to know its location. I have it in the AHK file as:
vValue_Setting_NPPlusPlus_Location := "C:\Program Files\Notepad++\notepad++.exe"

You can change it but then you also need to recompile the exe file, or you could use the AHK file (if AHK is installed on your system).
