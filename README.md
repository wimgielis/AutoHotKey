# AutoHotKey
A set of tools written in the automation language AutoHotKey

Download the exe file and store it anywhere you like.

A set of tools can be started by pressing the F10 function key. A menu will appear.
Most tools are to be used in Notepad++ or TM1 Architect/Perspectives (notably, Turbo Integrator and the Rules Editor).

The corresponding AHK file contains the source code. Please keep my headers inside. The original idea comes from here, although a lot of the code has changed and is created by myself: https://code.cubewise.com/ti-helper
To reload the menu, press Ctrl-F10.

I also put out my custom "user defined languages" in Notepad++. I add TM1 and AHK language syntax support.
The code in the exe file will be toggle it on automatically with some of the tools.

As Notepad++ is used and automated, we need to know its location. I have it in the AHK file as:
vValue_Setting_NPPlusPlus_Location := "C:\Program Files\Notepad++\notepad++.exe"

You can change it but then you also need to recompile the exe file, or you could use the AHK file (if AHK is installed on your system).

Last update:
- Pressing F10 is also allowed in Windows File Explorer (to cater for the tool described below)
- Jan. 30, 2021: a new TM1 model (in File Explorer, navigate to the desired folder and hit F10 and choose the new option in the menu)
- Jan. 31, 2021: automatically open up a PRO file from disk in Notepad++ based on: opened TI process in the TI editor, or selected text in Notepad++, or the selected error log file in File Explorer (enter in the AHK file the TM1 data directory: cPath_TM1_model_Main := "D:\path to my data directory with trailing backslash\" and use that one, optionally compile the AHK file to a new EXE file)
- Jan. 31, 2021: automatically backup with a timestamp what you are working on: a file in Notepad++, the TI process that you are looking at in TM1 Architect/Perspectives, the RUX file that you are looking at in TM1 Architect/Perspectives, even the Excel file you are looking at.
- Mar. 26, 2021: extensions to the functionality of "Generate AsciiOutput". You can select the variables in a certain order, which is then retained in the output
