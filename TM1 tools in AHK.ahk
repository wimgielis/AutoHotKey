; ####################################################
; #
; # Author: Wim Gielis
; # Contact: wim.gielis@gmail.com
; # URL: https://github.com/wimgielis?tab=repositories
; # Website: https://www.wimgielis.com
; # Date: January 30, 2021
; # Based on the TI helper AutoHotKey script: https://code.cubewise.com/ti-helper
; #
; ####################################################

; 1. in Turbo Integrator editor, Rules editor, Notepad++
;    a) F10: create an additional menu next to the mouse pointer
;    b) Ctrl-F10: reload the different menu's (used with F10)
; 2. in Turbo Integrator editor, Rules editor
;    a) Ctrl-k: comment code
;    b) Ctrl-Shift-k: uncomment code
;    c) Ctrl-l: delete the currently selected line (multiple lines default to the first line)
;    d) Ctrl-d: duplicate the currently selected line (multiple lines default to the first line)

#Persistent
#SingleInstance Force   ; Only opens a single instance of the script so reopening won't cause multiple instances
#NoEnv                  ; Recommended for performance and compatibility with future AutoHotkey releases
SendMode Input          ; Recommended for new scripts due to its superior speed and reliability
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory
CoordMode, Mouse, Client

;~~~~~~~~~~~~~~~~~~~
; GLOBAL OBJECTS
;~~~~~~~~~~~~~~~~~~~
ToolsArray := Object()

;~~~~~~~~~~~~~~~~~~~
; INITIALISATION
;~~~~~~~~~~~~~~~~~~~
cPath_TM1_model_Main := "D:\path to my data directory with trailing backslash\"
GoSub, LoadMenu

;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;   Additional shortcuts in TI and rules, mimicking Notepad++ shortcuts:
;   - Ctrl-k: comment code
;   - Ctrl-Shift-k: uncomment code
;   - Ctrl-l: delete the currently selected line (multiple lines default to the first line)
;   - Ctrl-d: duplicate the currently selected line (multiple lines default to the first line)
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

commenting:
#if WinActive("Turbo Integrator:") or WinActive("Rules Editor:")
^k:: ; <-- (TI, rules) indent code
{
    NewClip =
    clipboard =
    Send ^c
    ClipWait, 1

    if ErrorLevel <> 0
        return

    AutoTrim, off

    for each, Line in StrSplit(Clipboard, "`n", "`r")
    {
       if ( Line = "" )
          NewClip .= Line "`r`n"
       else
          NewClip .= "# " Line "`r`n"
    }
    StringTrimRight, NewClip, NewClip, 2

    clipboard = %NewClip%
    Send ^v
    Return
}

uncommenting:
^+k:: ; <-- (TI, rules) unindent code
{
    NewClip =
    clipboard =
    Send ^c
    ClipWait, 1

    if ErrorLevel <> 0
        return

    WaitingForFirstCharOfLine = y
    AutoTrim, off

    for each, Line in StrSplit(Clipboard, "`n", "`r")
    {
       StringTrimLeft, Line, Line, 2
       NewClip .= Line "`r`n"
    }
    StringTrimRight, NewClip, NewClip, 2

    clipboard = %NewClip%
    Send ^v
    Return
}

^l:: ; <-- (TI, rules) delete the current line
{
    SendInput, {Home}
    SendInput, {Shift down}{End}{Shift up}
    SendInput, {Delete}
    SendInput, {Backspace}
return
}

^d:: ; <-- (TI, rules) duplicate the current line
{
    SendInput, {Home}
    SendInput, {Shift down}{End}{Shift up}
    SendInput, ^c
    SendInput, {Home}
    SendInput, {Enter}
    SendInput, {Up}
    SendInput, ^v
    SendInput, {Right}
return
}


;~~~~~~~~~~~~~~~~~~~
; HOTKEYS (Note: there are more hotkeys above)
;~~~~~~~~~~~~~~~~~~~
#IfWinActive ahk_class ExploreWClass
^F10:: ; <-- (Notepad++) reload
    Reload
F10:: ;{ <--  ~ (Notepad++) Main menu ~
    Menu,menu,show
Return

#IfWinActive ahk_class CabinetWClass
^F10:: ; <-- (Notepad++) reload
    Reload
F10:: ;{ <--  ~ (Notepad++) Main menu ~
    Menu,menu,show
Return


#IfWinActive, ahk_class Notepad++
^F10:: ; <-- (Notepad++) reload
    Reload
F10:: ;{ <--  ~ (Notepad++) Main menu ~
    Menu,menu,show
Return

#IfWinActive, ahk_class AfxFrameOrView100u
^F10:: ; <-- (TI) reload
    Reload

; to resize the columns of the VARIABLES AND PARAMETERS window
F10:: ;{ <--  ~ (TI) Main menu ~
{
   WinGet, hWnd, ID, A
   SendMessage, 0x130B,,, SysTabControl327, % "ahk_id " hWnd ;TCM_GETCURSEL := 0x130B
   vSelection := ErrorLevel+1

   if (vSelection = 2)
   {
      ; resize the columns
      WinGetTitle, Window_Title, A
      MouseClickDrag, L, 1335, 75, 1140, 75
      MouseClickDrag, L, 1015, 75, 900, 75
      MouseClickDrag, L, 700, 75, 470, 75
      MouseClickDrag, L, 380, 75, 250, 75
      return
   }
   else if (vSelection = 4)
   {
      ; the Advanced tab

      WinGet, hWnd, ID, A
      SendMessage, 0x130B,,, SysTabControl321, % "ahk_id " hWnd ;TCM_GETCURSEL := 0x130B
      vSelection := ErrorLevel+1

      if (vSelection = 1)
      {
         ; resize the columns
         WinGetTitle, Window_Title, A
         MouseClickDrag, L, 437, 100, 850, 100
         MouseClickDrag, L, 337, 100, 380, 100
         MouseClickDrag, L, 237, 100, 215, 100
         MouseClickDrag, L, 137, 100, 210, 100
         return
      }
      else
      {
         ; Prolog - Metadata - Data - Epilog
         Menu,menu,show
         return
      }
   }
}
Return

#IfWinActive, ahk_class #32770
; to increase the size of the PARAMETERS window when you EXECUTE a process
F10:: ;{ <--  ~ (TI) run process, format the screen ~
{
   WinGetTitle, Window_Title, A
   WinMaximize

   ; resize the first column
   MouseClickDrag, L, 160, 40, 800, 40
   ; a double click in the first parameter value
   Click, 810, 55, 2
   return
}
Return


;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
;   LoadMenu function
;   This function loads menu items
;~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

LoadMenu:

    ; tools
	toolName:= "Code to Notepad"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	toolName:= "Generate AsciiOutput"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	toolName:= "Update PRO line numbers"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	; a separator
	toolName:=
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	toolName:= "Commenting"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	toolName:= "Uncommenting"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	; a separator
	toolName:=
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	toolName:= "Convert to Expand"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	toolName:= "Convert to | delimited"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	; a separator
	toolName:=
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	toolName:= "A new TM1 model (in File Explorer)"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	toolName:= "Open the PRO file"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	toolName:= "File backup with timestamp"

	ToolsArray.Insert([toolName])
	Menu,ToolsSubmenu,add,%toolName%,menuHandler

	menuName:="Tools"
	menuHandler:="ToolsSubmenu"

	Menu,menu,add,%menuName%,:%menuHandler%

Return

menuHandler:

    menuItem = %A_ThisMenuItem%

    nFound:=0

    if(menuItem = "Code to Notepad")
    {
        WinGetTitle, Window_Title, A
        if InStr( Window_Title, "Turbo Integrator:", false ) = 0
        {
            msgbox Please activate Turbo Integrator
            return
        }

        vText = % Get_TI_Code_Tab()

        ; get or create a Notepad++ session
        IfWinExist ahk_class Notepad++
        {
           WinActivate
           WinWaitActive
           Sleep 300
           WinMenuSelectItem, , , File, New ; open a new tab
        }
        Else
        {
           vValue_Setting_NPPlusPlus_Location := "C:\Program Files\Notepad++\notepad++.exe"
           IfExist, %vValue_Setting_NPPlusPlus_Location%
           {
               Run, %vValue_Setting_NPPlusPlus_Location%
               WinActivate ahk_class Notepad++
               WinWaitActive ahk_class Notepad++
               Sleep 300
               WinMenuSelectItem, , , File, New ; open a new tab
           }
           Else
           {
              try
                 {
                 Run notepad++
                 WinActivate ahk_class Notepad++
                 WinWaitActive ahk_class Notepad++
                 Sleep 300
                 WinMenuSelectItem, , , File, New ; open a new tab
                 }
              catch
                 {
                 Run notepad
                 }
           }
        }

        Clipboard := vText
        Send, ^v

        ; tab to space
        WinMenuSelectItem, , , 2&, 16&, 6&

        ; remove trailing spaces
        WinMenuSelectItem, , , 2&, 16&, 1&

        ; apply the TM1 language
        WinMenuSelectItem, , , Language, TM1

        Send ^{home}

        if ErrorLevel   ; i.e. it's not blank or zero.
            MsgBox, ERROR

        return
    }
    else if(menuItem = "Commenting")
    {
        Send {Esc 2}
        Gosub commenting
        return
    }
    else if(menuItem = "Uncommenting")
    {
        Send {Esc 2}
        Gosub uncommenting
        return
    }
    else if(menuItem = "Generate AsciiOutput")
    {

		; ####################################################
		; # Purpose of the script:
		; # - open up any (valid) PRO file in Notepad++ or TI process editor, or just a snippet of TI code without a regular PRO file
		; # - launch the menu and the gui
		; # - list all variables of a TI process
		; # - allow to generate the TI code for selected variables
		; # - it will be pasted at the cursor
		; # - Order of the selected (clicked) variables could be respected and AsciiOutput in that order
		; # - several other options are built-in
		; #
		; # To be added later:
		; # - Order of the selected variables could be respected and AsciiOutput in that order
		; # - A tab control on the gui containing a 'live' preview of what the code would create (with manual adjustment)
		; # - Use the Windows registry to save to and load from for choices in the gui (not needed to redo everything)
		; # - No REST API used and needed for now, I parse a text file and that's it
		; # - 
		; ####################################################


		vProcess := Get_Process_FullFileName( "", "___Current process" )
		if vProcess =
		   vProcess := "No PRO file"

		Gui, Destroy

		Gui, Add, Text, vFullFilename Disabled, %vProcess%

		Gui, Add, Button, gAO_SelectALL x15, Select all
		Gui, Add, Button, gAO_ReloadVars x+30, Reload variables
		Gui, Add, edit, vEdit_Select_Matching_Variables x+20 w80 r1 gSelect_Matching_Variables
		Gui, Add, Button, gAO_RemoveDups x+55, Remove duplicates
		Gui, Add, Button, gAO_ReIndex x+40, Re-index order
		
		Gui, Add, ListView, x10 w400 h500 gMyListView AltSubmit, Variable Type|Variable Name|Data Type|Order

		Gui, Add, ListBox, 8 x+10 w150 r9 Choose1 vShow_Which_Variables gShow_Which_Variables, All|Parameters|Data source variables|Data source new variables|Custom variables|   in the Prolog tab|   in the Metadata tab|   in the Data tab|   in the Epilog tab
		Gui, Add, ListBox, AltSubmit ReadOnly y+30 w150 r3 vSwitch_Text_Number gSwitch_Text_Number, Text <--> Number|--> Text|--> Number
		Gui, Add, Button, gAO_OK y+40, Generate output

		Gui, Add, ListBox, x10 w100 Choose1 vMyListBox_OutputFunction, AsciiOutput|TextOutput|LogOutput
		Gui, Add, ListBox, AltSubmit x+20 w100 Choose1 vMyListBox_OutputFolder, DataDir|LoggingDir|DebugDir
		Gui, Add, edit, vEdi1 x+20 w160 r1, test.txt
		Gui, Add, CheckBox, y+10 vVariable_Filename, Variable

        Gui, Add, CheckBox, x10 y+15 Checked vOutput_in_Selected_Order gOutput_in_Selected_Order, Output variables in selected order
        Gui, Add, CheckBox, x10 vAdd_Output_Settings, Add output settings
																				 

        Gui, Add, CheckBox, x10 vNames_of_variables, Insert names of variables
        Gui, Add, CheckBox, x10 vHeader, Add code for a header
        Gui, Add, CheckBox, x10 vNumbers_NumberFormat, Numbers get a numberformat
		Gui, Add, CheckBox, x25 y+5 vNumberToString_Enough, NumberToString suffices

		Gui, Add, ListBox, AltSubmit Choose1 x10 y+15 w150 r3 vLine_Splitter, No split|Multiline output|Split long lines

		Gui, Add, Edit, x250 y650 w50 Center
		Gui, Add, UpDown, Left vUpDown_IF Range0-10, 0
		Gui, Add, Text, x+5 yp+3, Add this many filters

		Gui, Add, Edit, x250 y685 w50 Center
		Gui, Add, UpDown, Left vUpDown_Repeat gUpDown_Repeat Range1-20, 1
		Gui, Add, Text, x+5 yp+3, Repeat output this many times
        Gui, Add, CheckBox, yp+20 vAdd_Area, Include area designation

        Gosub Read_in_AO_vars

		LV_ModifyCol()  ; Auto-size each column to fit its contents
        LV_ModifyCol(4, "45 Integer Center")
		Gui, Show, w600 h850, Output variables to a text file                                                         (c) 2021 - Wim Gielis
        return

        UpDown_Repeat:
        gui, Submit, noHide
		If UpDown_Repeat > 1
		    GuiControl,,Add_Area, 1
		else
		    GuiControl,,Add_Area, 0
		return
		
        Select_Matching_Variables:
		GuiControlGet, Edit_Select_Matching_Variables
		If( Edit_Select_Matching_Variables != "" )
		{
            LV_Modify(0, "-Select")
			Loop % LV_GetCount()
            {
                LV_GetText(Text, A_Index, 2)
                if ( Instr( Text, Edit_Select_Matching_Variables ) > 0 )
                    LV_Modify(A_Index, "+Select")
            }
        }	
		return
		
        MyListView:
        gui, Submit, noHide
		if ( A_GuiEvent = "Normal" )
        {
           if ( Output_in_Selected_Order = True )
           {
		       ; find the highest order number and increment
               Selected_Row = %A_EventInfo%
               dMax := 0
			   Loop % LV_GetCount()
               {
                   LV_GetText(Text, A_Index, 4)
                   Number := ("0" . Text), Number += 0
                   if ( Number > dMax )
                       dMax := Number
               }
               LV_GetText(Text, Selected_Row, 4)
               Number := ("0" . Text), Number += 0
		       If ( Number < dMax Or Number = 0 )
			       {
				   n := dMax + 1
		           LV_Modify(Selected_Row, "Col4", n)
				   }
		       ; how many have a number ?
		       If ( Number > 0 )
               {
			       countof := 0
			       Loop % LV_GetCount()
                   {
                       LV_GetText(Text, A_Index, 4)
                       If ( Text != "" )
                           countof += 1
                   }
			       If ( n > countof )
                   {
			           Loop % LV_GetCount()
                       {
                           LV_GetText(Text, A_Index, 4)
                           If ( Text != "" )
                           {
                               Number2 := ("0" . Text), Number2 += 0
                               If ( Number2 > Number )
				    		       LV_Modify(A_Index, "Col4", Number2 - 1)
                           }
                       }
                   }
               }
		   }

        }
		if ( A_GuiEvent = "RightClick" )
        {
		    ; clear the order for the selected row
		    LV_Modify(A_EventInfo, "Col4", "")
		}
		return

		Show_Which_Variables:
		gosub Read_in_AO_vars
		return

		Read_in_AO_vars:
		Gui, Submit, NoHide
		LV_Delete()

		guiControlGet, sFullFilename,, FullFilename
		if ( sFullFilename = "No PRO file" )
            {
			if ( sSaved_Title != "" )
			    Title := sSaved_Title
			else
				WinGetTitle, Title, A
			WinGetText, text, %Title%
            FileContents := text
			if ( sSaved_Title = "" )
			     sSaved_Title := Title
			}
		else
		    FileRead, FileContents, %sFullFilename%


		; make sure we have CRLF as the end of line character
        If ( Instr( FileContents, "`r`n" ) = 0 )
           If ( Instr( FileContents, "`n" ) > 0 )
		      FileContents := StrReplace(FileContents, "`n", "`r`n")
           else if ( Instr( FileContents, "`r" ) > 0 )
		      FileContents := StrReplace(FileContents, "`r", "`r`n")

		; 1. parameters
		if InStr( Show_Which_Variables, "All", false ) = 1
		Or InStr( Show_Which_Variables, "Parameters", false ) > 0
		{
			PRO_Text := StringBetween( FileContents, "560,", "590," )
			Nr_of_vars := ( CountLines( PRO_Text ) - 3 ) / 2
			If ( Nr_of_vars > 0 )
			{
				StringSplit, ParamArray, PRO_Text, `n
				Loop, %ParamArray0%
				{
					if A_index = 1
						Continue

					Index_type := A_Index + ( Floor(Nr_of_vars) * 1 ) + 1

					my_param_name := ParamArray%A_Index%
					; remove the linebreak
					StringTrimRight, my_param_name, my_param_name, 1

					my_param_type := ParamArray%Index_type%
					; remove the linebreak
					StringTrimRight, my_param_type, my_param_type, 1

					If ( my_param_type = 2 )
						LV_Add("", "Parameter", my_param_name, "Text")
					Else
						LV_Add("", "Parameter", my_param_name, "Number")

					if A_index > % Nr_of_vars
						Break
				}
			}
        }

        ; 2. data source variables
		if InStr( Show_Which_Variables, "All", false ) = 1
		Or InStr( Show_Which_Variables, "Data source variables", false ) > 0
		{
			PRO_Text := StringBetween( FileContents, "577,", "579," )
			Nr_of_vars := ( CountLines( PRO_Text ) - 3 ) / 2

			If ( Substr( StringBetween( FileContents, "562,", "586," ), 2, 4 ) = "VIEW" )
				Nr_of_vars_without_builtin := Nr_of_vars - 3
			Else
				Nr_of_vars_without_builtin := Nr_of_vars

			If ( Nr_of_vars > 0 )
			{
				StringSplit, ParamArray, PRO_Text, `n
				Loop, %ParamArray0%
				{
					if A_index = 1
						Continue

					Index_type := A_Index + ( Floor(Nr_of_vars) * 1 ) + 1

					my_param_name := ParamArray%A_Index%
					; remove the linebreak
					StringTrimRight, my_param_name, my_param_name, 1

					my_param_type := ParamArray%Index_type%
					; remove the linebreak
					StringTrimRight, my_param_type, my_param_type, 1

					If (( my_param_name = "NValue" ) Or ( my_param_name = "SValue" ) Or ( my_param_name = "value_is_STRING" ))
					{
						If ( my_param_type = 2 )
							LV_Add("", "View var", my_param_name, "Text")
						Else
							LV_Add("", "View var", my_param_name, "Number")
					}
					Else
					{
						If ( my_param_type = 2 )
							LV_Add("", "Data source var", my_param_name, "Text")
						Else
							LV_Add("", "Data source var", my_param_name, "Number")
					}

					if A_index > % Nr_of_vars
						Break
				}
            }
        }

		; 3. data source variables: newly created by the developer
		if InStr( Show_Which_Variables, "All", false ) = 1
		Or InStr( Show_Which_Variables, "Data source new variables", false ) > 0
		{
			PRO_Text := StringBetween( FileContents, "582,", "603," )
			Nr_of_vars := ( CountLines( PRO_Text ) - 2 ) / 1

			If ( Nr_of_vars > Nr_of_vars_without_builtin )
			{
				VarArray := StrSplit(PRO_Text, "`r`n")
				Loop % VarArray.MaxIndex()
				{
					if A_index - 1 <= Nr_of_vars_without_builtin
						Continue

					my_var_line := VarArray[A_Index]

					if ( InStr( my_var_line, "VarName", false ) = 1 )
					{
					   VarArray_FF := StrSplit(my_var_line, Chr(12))

					   my_var_name_full := VarArray_FF[1]
					   VarArray_EqualSign := StrSplit(my_var_name_full, "=")
					   my_var_name := VarArray_EqualSign[2]

					   my_var_type_full := VarArray_FF[2]
					   VarArray_EqualSign := StrSplit(my_var_type_full, "=")
					   my_var_type := VarArray_EqualSign[2]

					   my_var_coltype_full := VarArray_FF[3]
					   VarArray_EqualSign := StrSplit(my_var_coltype_full, "=")
					   my_var_coltype := VarArray_EqualSign[2]

					   If ( my_var_coltype <> 1165 )
					   {
						If ( my_var_type = 32 )
							LV_Add("", "Data source new var", my_var_name, "Text")
						Else
							LV_Add("", "Data source new var", my_var_name, "Number")
					   }
					}
				}
			}
		}

		; 4. custom variables created by the developer anywhere in the process code
		if InStr( Show_Which_Variables, "All", false ) = 1
		Or InStr( Show_Which_Variables, "Custom variables", false ) > 0
		Or InStr( Show_Which_Variables, "   in the Prolog tab", false ) > 0
		Or InStr( Show_Which_Variables, "   in the Metadata tab", false ) > 0
		Or InStr( Show_Which_Variables, "   in the Data tab", false ) > 0
		Or InStr( Show_Which_Variables, "   in the Epilog tab", false ) > 0
		{
			Text_of_vars :=
			if ( sFullFilename = "No PRO file" )
			{
				Text_of_vars_1234 := FileContents
				Text_of_vars_1    :=
				Text_of_vars_2    :=
				Text_of_vars_3    :=
				Text_of_vars_4    :=
			}
			else
			{
				Text_of_vars_1234 := StringBetween( FileContents, "572,", "576," )
				Text_of_vars_1    := StringBetween( FileContents, "572,", "573," )
				Text_of_vars_2    := StringBetween( FileContents, "573,", "574," )
				Text_of_vars_3    := StringBetween( FileContents, "574,", "575," )
				Text_of_vars_4    := StringBetween( FileContents, "575,", "576," )
			}

			if InStr( Show_Which_Variables, "All", false ) = 1
			Or InStr( Show_Which_Variables, "Custom variables", false ) > 0
				Text_of_vars := Text_of_vars_1234
			if InStr( Show_Which_Variables, "   in the Prolog tab", false ) > 0
				Text_of_vars .= Text_of_vars_1
			if InStr( Show_Which_Variables, "   in the Metadata tab", false ) > 0
				Text_of_vars .= Text_of_vars_2
			if InStr( Show_Which_Variables, "   in the Data tab", false ) > 0
				Text_of_vars .= Text_of_vars_3
			if InStr( Show_Which_Variables, "   in the Epilog tab", false ) > 0
				Text_of_vars .= Text_of_vars_4

			; search for variables
			allMatches := ""
			out := ""
			Pos = 1
			While Pos := RegExMatch(Text_of_vars, "(?<=\v)\s*?\w+?\s*?=", m, Pos+StrLen(m))
			{
			   m := StrReplace(m, "=", "")
			   m := StrReplace(m, A_Space, "")
               m := StrReplace(m, A_Tab, "")
			   m := StrReplace(m, "`r`n", "")
			   allMatches .= (!allMatches ? "" : "`n") m
			}

			Loop, parse, allMatches, `n
			If not RegExMatch(out, "\b" A_LoopField "\b") {
				out .= out ? "`n" : ""
				out .= A_LoopField
			}

			; to ease searching for the variable name in the code, then retrieve the following word/characters
			Text_of_vars_no_whitespace := StrReplace(Text_of_vars, A_Space)
			Text_of_vars_no_whitespace := StrReplace(Text_of_vars_no_whitespace, A_Tab)
			e := "`r`n"

            Lookup_VarType := ["2|'", "1|CellGetN", "2|CellGetS", "1|AttrN", "2|AttrS", "1|StringToNumber", "2|NumberToString", "2|SubsetGetElementName", "1|SubsetGetSize", "2|Trim", "2|Subst", "1|Scan", "1|ElparN", "2|Elpar", "1|ElIspar", "1|ElcompN", "2|ElComp", "1|ElIsComp", "2|Long", "1|Dimix", "1|Dimsiz", "2|Dimnm", "2|Dtype", "2|Expand", "2|Tabdim", "2|Delet", "2|Insrt", "2|Upper", "2|Lower", "2|Char", "2|DimensionElementPrincipalName", "2|GetProcessName", "2|TM1User", "2|Timst", "1|Dayno", "2|GetProcessErrorFileDirectory", "1|ExecuteProcess", "1|RunProcess" ]

			VarArray := StrSplit(out, "`n")
			Loop % VarArray.MaxIndex()
			{
				my_var_name := VarArray[A_Index]
				StringLower, vmy_var_name, my_var_name

				Found := 0

				if( vmy_var_name = "datasourceasciidelimiter"
				 Or vmy_var_name = "datasourceasciiquotecharacter"
				 Or vmy_var_name = "datasourceasciidecimalseparator"
				 Or vmy_var_name = "datasourceasciithousandseparator"
				 Or vmy_var_name = "datasourceasciiheaderrecords"
				 Or vmy_var_name = "datasourcetype"
				 Or vmy_var_name = "datasourcenameforserver"
				 Or vmy_var_name = "datasourcenameforclient"
				 Or vmy_var_name = "datasourcecubeview"
				 Or vmy_var_name = "datasourcedimensionsubset"
				 Or vmy_var_name = "datasourceusername"
				 Or vmy_var_name = "datasourcepassword"
				 Or vmy_var_name = "datasourcequery"
				 Or vmy_var_name = "vmask"
				 Or vmy_var_name = "vdec"
				 Or vmy_var_name = "v1000" )
				   	{

					   Found := 1
				   	   LV_Add("", "Custom variable", my_var_name, "Text")
                    }

				If( Found = 0 )
				{
                   for i, element in Lookup_VarType
                   {
                      Element_Type := SubStr(element, 1, InStr(element, "|") - 1) 
                      If( Element_Type == "1" )
				         Element_Type_Full := "Number"
                      Else
                         Element_Type_Full := "Text"
				   
                      Element_Function := SubStr(element, InStr(element, "|") + 1)
				   
                      if( InStr(Text_of_vars_no_whitespace, e . my_var_name . "=" . Element_Function ) > 0 )
                      {
				   	     Found := 1
				   	     LV_Add("", "Custom variable", my_var_name, Element_Type_Full)
					
				   		 Break
				   	  }
                   }
                }
				
				If( Found = 0 )
				{
				   Needle := e . my_var_name . "="
				   FoundPos := InStr(Text_of_vars_no_whitespace, Needle )
				   Following_Character := SubStr(Text_of_vars_no_whitespace, FoundPos + StrLen(Needle), 1)
				   if Following_Character Is Integer
                   {
                      Found := 1
                      LV_Add("", "Custom variable", my_var_name, "Number")
                   }
				}
				
				If( Found = 0 )
				{
				   vFirst_Letter := Substr( my_var_name, 1, 1 )
                   StringLower, vFirst_Letter, vFirst_Letter
                   if( vFirst_Letter = "n" )
                   {
                      Found := 1
                      LV_Add("", "Custom variable", my_var_name, "Number")
                   }
                   else if( vFirst_Letter = "s" )
                   {
                      Found := 1
                      LV_Add("", "Custom variable", my_var_name, "Text")
                   }
				}

				If( Found = 0 )
					LV_Add("", "Custom variable", my_var_name, "Text/Number")

			}
		}
		return

		Output_in_Selected_Order:
		Gui, Submit, NoHide		
		if ( Output_in_Selected_Order = True )
        {}
		else
		{
		    Loop % LV_GetCount()
                LV_Modify(A_Index, "Col4", "")
        }
        AO_ReIndex:
		; 1. read the order numbers and row numbers in an associative array
        oArray := {}
		Loop % LV_GetCount()
		{
           LV_GetText(my_order, A_Index, 4)
		   if ( my_order != "" )
              oArray[A_Index] := my_order
		}

		; 2. sort the array on the order numbers
        temp := {}
        for key, val in oArray
           temp[val] ? temp[val].Insert(key) : temp[val] := [key]
		
		; 3. loop over the sorted array and populate the order for the row number in ascending manner
		c := 0
        for i, x in temp
           for k, y in x
		   {
		      ; msgbox %i% - %k% - %y%
		      c += 1
              LV_Modify(y, "Col4", c)
           }

		return
		
		AO_RemoveDups:
		ControlList := "|"
		DupesList := 
		Loop % LV_GetCount()
		{
		  LV_GetText(RowText, A_Index, 2)
		  IfInString, ControlList, |%RowText%|
			DupesList = %A_Index%|%DupesList%
		  else
			ControlList = |%RowText%%ControlList%
		}
		; DupesList .= A_Index . "|"
        Loop, parse, DupesList, |
        {
           If A_LoopField is integer
              LV_Delete(A_LoopField)
        }	
		return
		AO_SelectALL:
		Gui, Submit, NoHide
		LV_Modify(0, "Select")

        if ( Output_in_Selected_Order = True )
		    Loop % LV_GetCount()
                LV_Modify(A_Index, "Col4", A_Index)
		return

		Switch_Text_Number:
		Gui, Submit, NoHide
		selectedCount := LV_GetCount("Selected")
		if not selectedCount
			{
			msgbox Please select 1 or more variables.
			Return
			}

		RowNumber := 0
		Loop
		{
			RowNumber := LV_GetNext(RowNumber)
			if not RowNumber
				break

			LV_GetText(Var_Type, RowNumber, 3)
			If ( Switch_Text_Number = 1 )
			{
				If( Var_Type == "Number" )
					LV_Modify(RowNumber, "Col3", "Text")
				else If( Var_Type == "Text" )
					LV_Modify(RowNumber, "Col3", "Number")
				else
					LV_Modify(RowNumber, "Col3", "Text")
			}
			Else If ( Switch_Text_Number = 2 )
			{
				If( Var_Type != "Text" )
					LV_Modify(RowNumber, "Col3", "Text")
				; Else
				; 	{
				; 	LV_GetText(Var_Name, RowNumber, 2)
				; 	msgbox %Var_Name% was al tekst
				; 	}
			}
			Else If ( Switch_Text_Number = 3 )
			{
				If( Var_Type != "Number" )
					LV_Modify(RowNumber, "Col3", "Number")
				;Else
				;	{
				;	LV_GetText(Var_Name, RowNumber, 2)
				;	msgbox %Var_Name% was al number
				;	}
			}
		}
		return

		AO_OK:
		Gui, Submit, NoHide
		e := "`r`n"
		sOutput :=
		m := Line_Splitter

		selectedCount := LV_GetCount("Selected")
		if not selectedCount
			{
			msgbox Please select 1 or more variables.
			Return
			}

		; output with settings
		If ( Variable_Filename = True )
		    {
			sOutput .= "vFile = "
			if MyListBox_OutputFolder = 1
					sOutput .= "'" . Edi1 . "';" . e
			else if MyListBox_OutputFolder = 2
				sOutput .= "GetProcessErrorFileDirectory | " . "'" . Edi1 . "';" . e
			else if MyListBox_OutputFolder = 3
				sOutput .= "'..\Debug\" . Edi1 . "';" . e
			}

		; output with settings
		If ( Add_Output_Settings = True )
		    {
			sOutput .= "DatasourceAsciiDelimiter = ';';" . e
			sOutput .= "DatasourceAsciiQuoteCharacter = '';" . e
			; sOutput .= "DatasourceAsciiDecimalSeparator = ',';" . e
			; sOutput .= "DatasourceAsciiThousandSeparator = '.';" . e
			}

		; output with numberformat
		If ( Numbers_NumberFormat = True )
			If (  NumberToString_Enough = False )
		    {
				sOutput .= "vMask = '#,0.##';" . e
				sOutput .= "vDec = ',';" . e
				sOutput .= "v1000 = '.';" . e
			}

		; output with header
		If ( Header = True )
		{
			sOutput .= "vCounter = 1;" . e
			sOutput .= "If( vCounter = 1 );" . e
			
			If ( MyListBox_OutputFunction = "LogOutput" )
				sOutput .= MyListBox_OutputFunction . "( 'INFO'"
			else
			{
				sOutput .= MyListBox_OutputFunction . "( "
				If ( Variable_Filename = True )
					sOutput .= "vFile"
				Else
				{
					if MyListBox_OutputFolder = 1
						sOutput .= "'" . Edi1 . "'"
					else if MyListBox_OutputFolder = 2
						sOutput .= " GetProcessErrorFileDirectory | " . "'" . Edi1 . "'"
					else if MyListBox_OutputFolder = 3
						sOutput .= "'..\Debug\" . Edi1 . "'"
				}
			}

			if ( Output_in_Selected_Order = True )
            {
		       LV_ModifyCol(4, "Sort")
			   Loop % LV_GetCount()
               {
                   LV_GetText(Text, A_Index, 4)
                   if ( Text != "" )
                   {
                       LV_GetText(Var_Name, A_Index, 2)
                       If ( Names_of_variables = False )
                          sOutput .= Add_Text(sOutput, ", '" . Var_Name . "'", m)
                       Else
                          sOutput .= Add_Text(sOutput, ", '" . Var_Name . "'" . ", '" . Var_Name . " value'", m)
                   }
               }
            }
            else
            {
		       RowNumber := 0
               Loop
               {
                  RowNumber := LV_GetNext(RowNumber)
                  if not RowNumber
                     break
                  LV_GetText(Var_Name, RowNumber, 2)
                  If ( Names_of_variables = False )
                     sOutput .= Add_Text(sOutput, ", '" . Var_Name . "'", m)
                  Else
                     sOutput .= Add_Text(sOutput, ", '" . Var_Name . "'" . ", '" . Var_Name . " value'", m)
               }
            }
			sOutput .= " );" . e
			sOutput .= "EndIf;" . e
			sOutput .= "vCounter = vCounter + 1;" . e . e
		}

		; output with run time values
		Loop % UpDown_Repeat
		{
			sOutput .= e
			Loop % UpDown_IF
				sOutput .= e . "If(  @= '' );"
			
			If ( MyListBox_OutputFunction = "LogOutput" )
				sOutput .= MyListBox_OutputFunction . "( 'INFO'"
			else
					  
		  
			{
				sOutput .= e . MyListBox_OutputFunction . "( "

				If ( Variable_Filename = True )
					sOutput .= "vFile"
				Else
				{
					if MyListBox_OutputFolder = 1
						sOutput .= "'" . Edi1 . "'"
					else if MyListBox_OutputFolder = 2
						sOutput .= " GetProcessErrorFileDirectory | " . "'" . Edi1 . "'"
					else if MyListBox_OutputFolder = 3
						sOutput .= "'..\Debug\" . Edi1 . "'"
				}
			}

			if ( Add_Area = True )
                   sOutput .= ", 'AREA_" . A_Index . "'"		       

			if ( Output_in_Selected_Order = True )
            {
			   LV_ModifyCol(4, "Sort")
			   Loop % LV_GetCount()
               {
                   LV_GetText(Text, A_Index, 4)
                   if ( Text != "" )
                   {
                        LV_GetText(Var_Name, A_Index, 2)
                        LV_GetText(Var_Type, A_Index, 3)
                        StringLower, vVar_Name, Var_Name
                        
                        If ( Names_of_variables = False )
                        {
                            If ( Var_Type = "Number" )
                            {
                                If ( Numbers_NumberFormat = False Or NumberToString_Enough = True Or vVar_Name = "value_is_string" )
                                    sOutput .= Add_Text(sOutput, ", NumberToString( " . Var_Name . " )", m)
                                Else
                                    sOutput .= Add_Text(sOutput, ", NumberToStringEx( " . Var_Name . ", vMask, vDec, v1000 )", m)
                            }
                            Else
                                sOutput .= Add_Text(sOutput, ", " . Var_Name, m)
                        }
                        Else
                        {
                            If ( Var_Type = "Number" )
                            {
                                If ( Numbers_NumberFormat = False Or NumberToString_Enough = True Or vVar_Name = "value_is_string" )
                                    sOutput .= Add_Text(sOutput, ", '" . Var_Name . ": ', " . "NumberToString( " . Var_Name . " )", m)
                                Else
                                    sOutput .= Add_Text(sOutput, ", '" . Var_Name . ": ', " . "NumberToStringEx( " . Var_Name . ", vMask, vDec, v1000 )", m)
                            }
                            Else
                                sOutput .= Add_Text(sOutput, ", '" . Var_Name . ": ', " . Var_Name, m)
                        }
                    }
               }
            }
			else
			{
                RowNumber := 0
                Loop
                {
                    RowNumber := LV_GetNext(RowNumber)
                    if not RowNumber
                        break
                    LV_GetText(Var_Name, RowNumber, 2)
                    LV_GetText(Var_Type, RowNumber, 3)
                    StringLower, vVar_Name, Var_Name
                    
                    If ( Names_of_variables = False )
                    {
                        If ( Var_Type = "Number" )
                        {
                            If ( Numbers_NumberFormat = False Or NumberToString_Enough = True Or vVar_Name = "value_is_string" )
                                sOutput .= Add_Text(sOutput, ", NumberToString( " . Var_Name . " )", m)
                            Else
                                sOutput .= Add_Text(sOutput, ", NumberToStringEx( " . Var_Name . ", vMask, vDec, v1000 )", m)
                        }
                        Else
                            sOutput .= Add_Text(sOutput, ", " . Var_Name, m)
                    }
                    Else
                    {
                        If ( Var_Type = "Number" )
                        {
                            If ( Numbers_NumberFormat = False Or NumberToString_Enough = True Or vVar_Name = "value_is_string" )
                                sOutput .= Add_Text(sOutput, ", '" . Var_Name . ": ', " . "NumberToString( " . Var_Name . " )", m)
                            Else
                                sOutput .= Add_Text(sOutput, ", '" . Var_Name . ": ', " . "NumberToStringEx( " . Var_Name . ", vMask, vDec, v1000 )", m)
                        }
                        Else
                            sOutput .= Add_Text(sOutput, ", '" . Var_Name . ": ', " . Var_Name, m)
                    }
                }
			}
			sOutput .= " );"

			Loop % UpDown_IF
				sOutput .= e . "EndIf;"

			sOutput .= e . e
		}


		Clipboard := sOutput
		Gui, Cancel
		Send ^v
		return

		AO_ReloadVars:
		gosub Read_in_AO_vars
        return
    }
    else if(menuItem = "File backup with timestamp")
    {
	
		; ####################################################
		; # Purpose of the script:
		; - automatically backup with a timestamp what you are working on
		; - this is done for:
		; - * a file in Notepad++
		; - * the PRO file that you are (implicitly) looking at in TM1 Architect/Perspectives in the TI editor
		; - * the RUX file that you are (implicitly) looking at in TM1 Architect/Perspectives in the Rules editor
		; - * an Excel file
		; ####################################################

		WinGetTitle, Window_Title, A
		FullCaption := Window_Title

		FormatTime, suffix,, (yyyy-MM-dd HHmmss)
		; suffix := " (" a_YYYY "-" a_MM "-" a_DD " " a_Hour a_Min a_Sec ")"

		vFile_New = 

		if InStr( Window_Title, "Turbo Integrator:", false ) = 1
		   {
		   StringSplit, process, FullCaption, "->"
		   process := trim(process%process0%)
		   vFile = %cPath_TM1_model_Main%%process%.pro

		   IfNotExist, %vFile%
			  {
			  Msgbox TI process %process% leads to a file that does not exist: %vFile%
			  return
			  }
		   }

		else if InStr( Window_Title, "Rules Editor:", false ) = 1
		   {
		   StringSplit, cube, FullCaption, "->"
		   cube := trim(cube%cube0%)
		   vFile = %cPath_TM1_model_Main%%cube%.rux
		   IfNotExist, %vFile%
			  {
			  Msgbox Rule for cube %cube% leads to a file that does not exist: %vFile%
			  return
			  }
		   }

		else if WinActive("ahk_class Notepad++")
		   {
		   WinMenuSelectItem, , , Edit, Copy to Clipboard, Current Full File path to Clipboard
		   Sleep 100
		   vFile := Clipboard
		   }

		else
		   {
		   vFile := % ComObjActive("excel.application").ActiveWorkbook.FullName
		   vFile := % fRework_Path(vFile)
		   }

		; copy the file with a timestamp
		IfExist, %vFile%
		   {
		   SplitPath, vFile, , dir, ext, f
		   vFile_New := % dir . "\" . f . " " suffix "." ext
		   }

		If vFile !=
		   If vFile_New !=
			  FileCopy, %vFile%, %vFile_New%
	}
        else if(menuItem = "Open the PRO file")
    {

		; ####################################################
		; # Purpose of the script:
		; - open up any (valid) PRO file in Notepad++
		; - this is done based on the following conditions:
		; - * if the clipboard contains a valid PRO file name in the TM1 data directory, this one is taken
		; - * if a TI process is opened up in the TI editor, this one is taken
		; - * if a TM1 processerror log file is selected in Windows File Explorer, the process name is parsed out and this one is taken
		; - * if still no luck, the selected text in Notepad++ is copied into the clipboard and then this one is taken
		; -     (so please select the name of a TI process, for example inside an ExecuteProcess statement)
		; ####################################################

        WinGetTitle, Window_Title, A
        WinGetClass, class, A
        FullCaption := Window_Title

        process := Clipboard
        vFile=%cPath_TM1_model_Main%%process%.pro
        IfExist, %vFile%
           { }
        else if InStr( Window_Title, "Turbo Integrator:", false ) = 1
           {
           ; FullCaption = %Window_Title%
           ; StringReplace, FullCaption, FullCaption, ->, ``, All
           ; StringSplit, arrprocess, FullCaption, ``
           ; process := trim(arrprocess2)
		   StringSplit, process, FullCaption, "->"
		   process := trim(process%process0%)
           }
        else if (class="CabinetWClass")||(class="ExploreWClass")||(class="Progman") ; File Explorer
           {
           vFullFile = % Explorer_GetSelection()
           SplitPath, vFullFile, , dir, extension, name_no_ext
           If extension = log
           {
           process := SubStr(name_no_ext, 41)
           }
           ; msgbox, % Explorer_GetSelection()
           }
        else
           {
           clipboard = ; clear the clipboard
           Send ^c
           ClipWait  ; Wait for the clipboard to contain text
           process := trim(clipboard)
           }

        vFile=%cPath_TM1_model_Main%%process%.pro
        IfExist, %vFile%
           {
           Run, "%vFile%"
           WinWait ahk_class Notepad++
           WinWaitActive ahk_class Notepad++
           WinMaximize
           }
        else
           msgbox %vFile% does not exist
    }

	else if(menuItem = "A new TM1 model (in File Explorer)")
    {
		; ####################################################
		; # Purpose of the script:
		; - Create a shortcut or service to run a TM1 model
		; - output is a shortcut or a service to run a TM1 model
		; - with a userform to guide the user
		; - when asked to browse to a folder, browse to the folder containing the tm1s.cfg file
		; - the tm1s.cfg file can be created/updated
		; - the config file for TM1 is not necessarily part of the TM1 data directory
		; - there is only 1 folder for the TM1 data directory
		; - the TM1 logging directory is part of the TM1 data directory, and is called "Logs", unless you choose the logging directory yourself
		; - TM1 server name can be set too. Do this before launching this tool. A number of other settings are set in this script, though, for your convenience
		; - you can also delete all commentary lines (lines starting with #). For instance, IBM's cfg file is quite big and not needed for simple models
		; ####################################################

; get settings, for example from an ini configuration file (now it's hardcoded here)
cPath_TM1s_exe := "C:\ibm\cognos\tm1_64\bin64\tm1s.exe"
cPath_TM1s_exe := ReplaceEnvVars(cPath_TM1s_exe)
IfNotExist, %cPath_TM1s_exe%
   cPath_TM1s_exe := "C:\ibm\cognos\tm1_64\bin64\tm1s.exe"

SplitPath, cPath_TM1s_exe, , dir
cPath_TM1sd_exe := % dir . "\tm1sd.exe"

currentpath :=
{
    v =
    v := GetSelectedFileFolderNames_UnderCursor()
	if v
	   currentpath = %v%
	else
	{
        v := GetSelectedFileFolderNames_FolderYouSee()
	    if v
	       currentpath = %v%
	    else
	       return
	}
}
currentpath := StrSplit(currentpath, "`n" )[1]

if InStr( FileExist(currentpath), "D") == 0
    {
	; take the folder
    SplitPath, currentpath, , currentpath
    }

gui, destroy
gui, tm1_as_app:new

gui, add, Button, vBut1 gBut x10 y+20 w150, Select folder of TM1s.cfg
gui, add, edit, vEdi1 x+20 w500 r1 gChanged_Editbox
gui, add, CheckBox, vChbx1 x+20

gui, add, Button, vBut2 gBut x10 w150, Select data dir
gui, add, edit, vEdi2 x+20 w500 r1 gChanged_Editbox
gui, add, CheckBox, vChbx2 x+20

gui, add, Button, vBut3 gBut x10 w150, Select logging dir
gui, add, edit, vEdi3 x+20 w500 r1 gChanged_Editbox
gui, add, CheckBox, vChbx3 x+20

gui, add, Button, vBut4 gBut x10 w150, TM1 server exe
gui, add, edit, vEdi4 x+20 w500 r1 gChanged_Editbox
gui, add, CheckBox, vChbx4 x+20

gui, add, Button, vBut6 gBut x10 w150, Select folder of shortcut
gui, add, edit, vEdi6 x+20 w500 r1 gChanged_Editbox
gui, add, CheckBox, vChbx6 x+20

gui, add, Text, x10 w150, Shortcut / service name
gui, add, edit, vEdi5 x+20 w500 r1 gChanged_Editbox
gui, add, CheckBox, vChbx5 x+20

gui, add, Text, x10 w150, Specify the TM1 model name
gui, add, edit, vEdi7 x+20 w500 r1 gChanged_Editbox
gui, add, CheckBox, vChbx7 x+20

gui, add, Text, x10 w150, Username for the service
gui, add, edit, vEdi8 x+20 w500 r1 gChanged_Editbox
gui, add, CheckBox, vChbx8 x+20

gui, add, Text, x10 w150, Password for the username
gui, add, edit, vEdi9 x+20 w500 r1 gChanged_Editbox Password
gui, add, CheckBox, vChbx9 x+20

Gui, add, ListBox, r3 x250 w200 choose1 gChanged_AppOrService vAppOrService, Create a shortcut|Create a service|Don't create shortcut nor service
Gui, add, CheckBox, Checked0 vChbx_RunModel x10 y+20, Run the TM1 model ?
Gui, add, CheckBox, Checked vChbx_Put_In_Default_Values, Put in default values ?
Gui, add, CheckBox, Checked0 x10 vChbx_Create_Addit_Folders, Create additional folders ?
Gui, add, CheckBox, Checked vChbx_Remove_Unnecessary_Lines, Remove unnecessary lines that are commented out ?
Gui, add, CheckBox, Checked0 vChbx_Delete_Client_Settings, Delete clients settings in the respective cub files ?
Gui, add, CheckBox, Checked0 vChbx_Delete_Feeders_Files, Delete feeders files ?
Gui, add, CheckBox, Checked0 vChbx_Delete_Controller_Objects, Delete Controller objects ?
Gui, add, CheckBox, Checked0 vChbx_Delete_Applications, Delete applications ?
Gui, add, CheckBox, Checked0 vChbx_Delete_DollarExtensionFiles, Delete dollar extension files ?
Gui, add, CheckBox, Checked0 vChbx_Delete_BackupFiles, Delete backup files ?
Gui, add, ListBox, r2 choose1 vAdminHostChoice, Localhost|%A_computername%
Gui, add, ListBox, r4 choose4 vChoresChoice, Delete chores|Deactivate chores|Activate chores|Leave untouched
gui, add, Button, vBut9 gBut x+150 w100, Temporary
gui, add, Button, vBut10 gBut x+150 w100, Folder nemen

GuiControl,,Edi1, %currentpath%

IfExist %currentpath%\TM1Data
   GuiControl,,Edi2, %currentpath%\TM1Data
else
   {
   IfExist %currentpath%\Data
	  GuiControl,,Edi2, %currentpath%\Data
   Else
      {
      p := % GetParentDir(currentpath, 1)
      IfExist %p%\TM1Data
         GuiControl,,Edi2, %p%\TM1Data
      else
         {
         IfExist %p%\Data
            GuiControl,,Edi2, %p%\Data
         ; else
            ; GuiControl,,Edi2, %currentpath%
         }
      }
   }


IfExist %currentpath%\TM1Logs
   GuiControl,,Edi3, %currentpath%\TM1Logs
else
   {
   IfExist %currentpath%\Logs
	  GuiControl,,Edi3, %currentpath%\Logs
   Else
      {
      p := % GetParentDir(currentpath, 1)
      IfExist %p%\TM1Logs
         GuiControl,,Edi3, %p%\TM1Logs
      else
         {
         IfExist %p%\Logs
            GuiControl,,Edi3, %p%\Logs
         ; else
            ; GuiControl,,Edi3, %currentpath%
         }
      }
   }


GuiControl,,Edi4, %cPath_TM1s_exe%
GuiControl,,Edi8,
GuiControl, Disable, Edi8
GuiControl, Disable, Chbx8
GuiControl,,Edi9,

GuiControl, Disable, Edi9
GuiControl, Disable, Chbx9

gui, add, Button, vBut7 gBut w100, Exit
gui, add, Button, vBut8 gBut x+150 w200 Default, Create

gui, show, , Run a TM1 model

return

But:
  gui, Submit, NoHide
  If (a_guicontrol = "But1") {
	FileSelectFolder, fld, *%Edi1%, 3, Select the folder containing the tm1s.cfg file`n(if missing it can be generated)
	If ErrorLevel
       {
       GuiControl,,Edi1,
	   }
    Else
	   {
	   fld := RegExReplace(fld, "\\$")
       GuiControl,,Edi1, % fld
	   }

  }
  If (a_guicontrol = "But2") {
    If Edi2 <>
       FileSelectFolder, fld, *%Edi2%
    Else
       FileSelectFolder, fld, *%Edi1%
    If ErrorLevel
       {
       GuiControl,,Edi2,
	   }
    Else
	   {
	   fld := RegExReplace(fld, "\\$")
       GuiControl,,Edi2, % fld
	   }
  }
  If (a_guicontrol = "But3") {
    If Edi3 <>
       FileSelectFolder, fld, *%Edi3%
    Else
       FileSelectFolder, fld, *%Edi1%
    If ErrorLevel
       {
       GuiControl,,Edi3,
	   }
    Else
	   {
	   fld := RegExReplace(fld, "\\$")
       GuiControl,,Edi3, % fld
	   }
  }
  If (a_guicontrol = "But4") {
	FileSelectFile, fil, 3, %cPath_TM1s_exe%, Select the TM1s.exe file (*.exe)
    If ErrorLevel
       {
       GuiControl,,Edi4,
	   }
    Else
	   {
       GuiControl,,Edi4, % fil
	   }
  }
  If (a_guicontrol = "But6") {
	FileSelectFolder, fld, *%Edi1%, 3, Select the folder that will contain the shortcut
	If ErrorLevel
       {
       GuiControl,,Edi6,
	   }
    Else
	   {
	   fld := RegExReplace(fld, "\\$")
       GuiControl,,Edi6, % fld
	   }

  }
  If (a_guicontrol = "But7") {
    gui, destroy
  }
  If (a_guicontrol = "But9") {
    msgbox Make sure the data directory is temporary\Data instead of temporary !
    GuiControl,,Edi1, C:\temporary
    GuiControl,,Edi2, C:\temporary\Data
    GuiControl,,Edi3, C:\temporary\Logs
    GuiControl,,Edi5, temporary
    GuiControl,,Edi6, C:\temporary
    GuiControl,,Edi7, temporary
	FileCreateDir, C:\temporary\Data
	FileCreateDir, C:\temporary\Logs
	GuiControl, , Chbx_RunModel, 1
  }
  If (a_guicontrol = "But10") {
    msgbox Make sure the data directory is %Edi1%\Data or please change it
    GuiControl,,Edi2, %Edi1%\Data
    GuiControl,,Edi3, %Edi1%\Logs

	splits := StrSplit(Edi1, "\")
	t := % splits[splits.MaxIndex()]
	; StringUpper, t, t
	; StringReplace, t, t, A_Space,   (deze code is fout)

    GuiControl,,Edi5, %t%
    GuiControl,,Edi6, %Edi1%
    GuiControl,,Edi7, %t%
	FileCreateDir, %Edi1%\Data
	FileCreateDir, %Edi1%\Logs
	GuiControl, , Chbx_RunModel, 1
  }
  If (a_guicontrol = "But8") {

    c01 = %Edi1%\tm1s.cfg

    ; treating the TM1 data directory
    vExisting_Cfg := FileExist(c01)
    IniWrite, %Edi7%, %c01%, TM1S, ServerName

    DataDirFolder := Edi2
    Folder_Data_Dir = %DataDirFolder%
    IniWrite, %Folder_Data_Dir%  , %c01%, TM1S, DataBaseDirectory

    LoggingDirFolder := Edi3
    Folder_Log_Dir = %LoggingDirFolder%
	If FileExist( Folder_Log_Dir )
       IniWrite, %Folder_Log_Dir%   , %c01%, TM1S, LoggingDirectory
	Else
	   {
          Folder_Log_Dir = %Folder_Data_Dir%\Logs
	      FileCreateDir, %Folder_Log_Dir%
          IniWrite, %Folder_Log_Dir%   , %c01%, TM1S, LoggingDirectory
	   }

	IniWrite, %AdminHostChoice%  , %c01%, TM1S, AdminHost

    ; add/update other much used settings in the TM1 data directory
    If Chbx_Put_In_Default_Values = 1
    {
    IniWrite, 1                  , %c01%, TM1S, IntegratedSecurityMode
    IniWrite, F                  , %c01%, TM1S, UseSSL
    IniWrite, T                  , %c01%, TM1S, EnableTIDebugging
    IniWrite, T                  , %c01%, TM1S, PersistentFeeders
    IniWrite, T                  , %c01%, TM1S, EnableNewHierarchyCreation

    ; OutputVar := Random( 8005, 8100 )
    Random, OutputVar, 8005, 8100
    If OutputVar = 8080
       Random, OutputVar, 8005, 8100
    IniWrite, %OutputVar%        , %c01%, TM1S, HttpPortNumber

    ; OutputVar := Random( 5000, 49151 )
    Random, OutputVar, 5000, 49151
    IniWrite, %OutputVar%        , %c01%, TM1S, PortNumber

    IniWrite, T                  , %c01%, TM1S, ParallelInteraction
    IniWrite, T                  , %c01%, TM1S, EnableTIDebugging
    IniWrite, T                  , %c01%, TM1S, ForceReevaluationOfFeedersForFedCellsOnDataChange
    IniWrite, all                , %c01%, TM1S, MTQ

    EnvGet, ProcessorCount, number_of_processors
    ProcessorCount--
    IniWrite, %ProcessorCount%   , %c01%, TM1S, MaximumCubeLoadThreads

    IniWrite, F                  , %c01%, TM1S, TopLogging
    IniWrite, 2                  , %c01%, TM1S, TopScanFrequency
    IniWrite, T                  , %c01%, TM1S, TopScanMode.Threads

    IniWrite, F                  , %c01%, TM1S, MTFeeders
    IniWrite, Kerberos           , %c01%, TM1S, SecurityPackageName
    IniWrite, ENG                , %c01%, TM1S, Language
    IniWrite, T                  , %c01%, TM1S, MTQQuery
    IniWrite, 5000               , %c01%, TM1S, MaximumViewSize
    IniWrite, T                  , %c01%, TM1S, VersionedListControlDimensions
    IniWrite, T                  , %c01%, TM1S, LoadPrivateSubsetsOnStartup
    IniWrite, T                  , %c01%, TM1S, ReduceCubeLockingOnDimensionUpdate
    IniWrite, T                  , %c01%, TM1S, EventLogging
    IniWrite, 1                  , %c01%, TM1S, EventScanFrequency
    IniWrite, 0                  , %c01%, TM1S, EventThreshold.PooledMemoryInMB
    IniWrite, 5                  , %c01%, TM1S, EventThreshold.ThreadBlockingNumber
    IniWrite, 600                , %c01%, TM1S, EventThreshold.ThreadRunningTime
    IniWrite, 20                 , %c01%, TM1S, EventThreshold.ThreadWaitingTime
    IniWrite, T                  , %c01%, TM1S, PullInvalidationSubsets
    IniWrite, F                  , %c01%, TM1S, PerformanceMonitorOn
    IniWrite, F                  , %c01%, TM1S, AuditLogOn
    IniWrite, 1800s              , %c01%, TM1S, AuditLogUpdateInterval
    IniWrite, 2 GB               , %c01%, TM1S, AuditLogMaxFileSize
    IniWrite, 1 GB               , %c01%, TM1S, AuditLogMaxQueryMemory
    IniWrite, 10                 , %c01%, TM1S, ClientPropertiesSyncInterval

    IniRead, LoggingDirectory    , %c01%, TM1S, LoggingDirectory,
    FileCreateDir, %LoggingDirectory%\FileRetry
    IniWrite, %LoggingDirectory%\FileRetry , %c01%, TM1S, FileRetry.FileSpec
    IniWrite, 5                  , %c01%, TM1S, FileRetry.Count
    IniWrite, 2000               , %c01%, TM1S, FileRetry.Delay

    IniRead, LoggingDirectory    , %c01%, TM1S, LoggingDirectory,
    FileCreateDir, %LoggingDirectory%\RawStore
    IniWrite, %LoggingDirectory%\RawStore , %c01%, TM1S, RawStoreDirectory
    FileCreateDir, %LoggingDirectory%\DistribPlanningOutputDir
    IniWrite, %LoggingDirectory%\DistribPlanningOutputDir , %c01%, TM1S, DistributedPlanningOutputDir

    IniDelete                    , %c01%, TM1S, IPAddressV4
    IniDelete                    , %c01%, TM1S, ServerCAMURI
    IniDelete                    , %c01%, TM1S, ClientCAMURI
    IniDelete                    , %c01%, TM1S, ClientPingCAMPassport

    ; for my own generated files in custom TI processes, could be useful if it already exists
    FileCreateDir, %LoggingDirectory%\TM1 output

	}

	; create additional folders
    If Chbx_Create_Addit_Folders = 1
        {
	    FileCreateDir, %Edi6%\Install files
	    FileCreateDir, %Edi6%\Project files\Control files
	    FileCreateDir, %Edi6%\Project files\Documentation
	    FileCreateDir, %Edi6%\Project files\Input\Data
	    FileCreateDir, %Edi6%\Project files\Input\Metadata
	    FileCreateDir, %Edi6%\Project files\Migration\From DEV to PROD
	    FileCreateDir, %Edi6%\Project files\Migration\From PROD to DEV
	    FileCreateDir, %Edi6%\Project files\Output
	    FileCreateDir, %Edi6%\Project files\Progress
	    FileCreateDir, %Edi6%\Project files\Templates
	    FileCreateDir, %Edi6%\Project files\Testing
	    FileCreateDir, %Edi6%\Scripts
	    FileCreateDir, %Edi6%\TM1Backup
	    }

	; delete unnecessary lines, only for already existing tm1s.cfg files
    If Chbx_Remove_Unnecessary_Lines = 1
        {
        If vExisting_Cfg
           {
           all =
           Loop, Read, %c01%
           {
           If A_LoopReadLine=
               Continue
	           StringLeft, tmp, A_LoopReadLine, 1
               If( tmp != "#" ) ;If this line isn't starting with # character
                   all = %all%`r`n%A_LoopReadLine%
           }
           StringTrimLeft, all, all, 2 ;Removes first `r`n line
           FileDelete, %c01%
           FileAppend, %all%, %c01%
           }
        }

    ; delete 4 cub files if the client settings need to be removed (CAM settings for example that wouldn't work anyway)
    If Chbx_Delete_Client_Settings = 1
        {
           ; FileDelete, %Folder_Data_Dir%\}Client*.cub
           FileDelete, %Folder_Data_Dir%\}ClientSettings.cub
           FileDelete, %Folder_Data_Dir%\}ClientCAMAssociatedGroups.cub
           FileDelete, %Folder_Data_Dir%\}ClientGroups.cub
           FileDelete, %Folder_Data_Dir%\}ClientProperties.cub
        }

    ; delete feeders files
    If Chbx_Delete_Feeders_Files = 1
        {
           FileDelete, %Folder_Data_Dir%\*.feeders
        }

    ; delete Controller cubes
    If Chbx_Delete_Controller_Objects = 1
        {
           FileDelete, %Folder_Data_Dir%\*Controllerfap*.cub
           FileDelete, %Folder_Data_Dir%\*Controllerfap*.feeders
           FileDelete, %Folder_Data_Dir%\*Controllerfap*.rux
           FileDelete, %Folder_Data_Dir%\*RPT_*.cub
           FileDelete, %Folder_Data_Dir%\*RPT_*.feeders
           FileDelete, %Folder_Data_Dir%\*RPT_*.rux
           FileDelete, %Folder_Data_Dir%\*RPT_*.pro
        }

    ; delete applications
    If Chbx_Delete_Applications = 1
        {
           FileRemoveDir, %Folder_Data_Dir%\}Applications, 1
           FileRemoveDir, %Folder_Data_Dir%\}Externals, 1
        }

    ; delete dollar extension files
    If Chbx_Delete_DollarExtensionFiles = 1
        {
           FileDelete, %Folder_Data_Dir%\*.*$
        }

    ; delete backup files
    If vChbx_Delete_BackupFiles = 1
        {
           FileDelete, %Folder_Data_Dir%\*.bak
           Loop, Files, %Folder_Data_Dir%\*.*
           {
             if RegExMatch( A_LoopFileExt, "\d{14}" )
               FileDelete, %A_LoopFileFullPath%
           }
        }

    ; chores logic
	If ChoresChoice = Delete chores
	   {
       IniDelete                    , %c01%, TM1S, StartupChores
	   FileDelete, %Folder_Data_Dir%\*.cho
	   }
    Else If ChoresChoice = Deactivate chores
	   {
           IniDelete                    , %c01%, TM1S, StartupChores
           Loop Files, %Folder_Data_Dir%\*.cho
           {
		   FileRead, OutputVar, %A_LoopFileFullPath%
           SearchTerm := "533,1"
		   If InStr(OutputVar, SearchTerm )
           {
           all =
           Loop, Read, %A_LoopFileFullPath%
           {
		   If A_LoopReadLine=
               Continue
	           StringLeft, tmp, A_LoopReadLine, 5
               If tmp = 533,1
				  all = %all%`r`n533,0
               Else
                  all = %all%`r`n%A_LoopReadLine%
           }
           StringTrimLeft, all, all, 2
           FileDelete, %A_LoopFileFullPath%
           FileAppend, %all%, %A_LoopFileFullPath%
           }
           }
       }
    Else If ChoresChoice = Activate chores
	   {
           Loop Files, %Folder_Data_Dir%\*.cho
           {
		   FileRead, OutputVar, %A_LoopFileFullPath%
           SearchTerm := "533,0"
		   If InStr(OutputVar, SearchTerm )
           {
           all =
           Loop, Read, %A_LoopFileFullPath%
           {
		   If A_LoopReadLine=
               Continue
	           StringLeft, tmp, A_LoopReadLine, 5
               If tmp = 533,0
				  all = %all%`r`n533,1
               Else
                  all = %all%`r`n%A_LoopReadLine%
           }
           StringTrimLeft, all, all, 2
           FileDelete, %A_LoopFileFullPath%
           FileAppend, %all%, %A_LoopFileFullPath%
           }
           }
       }

	If AppOrService = Don't create shortcut nor service
       {
	      msgbox User chose to not create a service or shortcut
	      Return
	   }

	else If AppOrService = Create a shortcut
       {
	   ; the shortcut to launch the model
       FileCreateShortcut, %cPath_TM1s_exe%, %Edi6%\%Edi5%.lnk, , -z"%Edi1%"
	   ; run the shortcut / TM1 model
       If Chbx_RunModel = 1
          Run, %Edi6%\%Edi5%.lnk
	   }
	Else If AppOrService = Create a service
       {
       ; full_command_line := GetParentDir( cPath_TM1s_exe, 1 )
       ; full_command_line := full_command_line . "\tm1sd.exe -install"
       full_command_line := cPath_TM1sd_exe
       full_command_line := full_command_line . " -n""" . Edi5 . """"
       full_command_line := full_command_line . " -z""" . GetParentDir( c01, 1) . """"
       full_command_line := full_command_line . " -u" . Edi8
       full_command_line := full_command_line . " -w" . Edi9

	   If (A_IsAdmin)
          RunWait %full_command_line%,, Hide
       Else
          RunWait *RunAs %full_command_line%,, Hide
       if (ErrorLevel = "ERROR")
          MsgBox An error was produced when creating the service to run the TM1 model. Please investigate.

       full_command_line := "sc config"
       full_command_line := full_command_line . " """ . Edi5 . """"
       full_command_line := full_command_line . " start=auto"
	   If (A_IsAdmin)
	      RunWait %full_command_line%,, Hide

	   ; run the shortcut / TM1 model
       If Chbx_RunModel = 1
          RunWait, cmd /c net start "%Edi5%",, Hide
	   }

    Return
  }
Return

Changed_AppOrService:
  gui, Submit, NoHide
  If AppOrService = Create a shortcut
     {
	 GuiControl, Enable, But6
	 GuiControl, Enable, Edi6
	 GuiControl, Enable, Chbx6

	 GuiControl, Disable, Edi8
	 GuiControl, Disable, Chbx8

	 GuiControl, Disable, Edi9
	 GuiControl, Disable, Chbx9

	 GuiControl,,Edi4, %cPath_TM1s_exe%
	 }
  Else If AppOrService = Create a service
     {
	 GuiControl, Disable, But6
	 GuiControl, Disable, Edi6
	 GuiControl, Disable, Chbx6

	 GuiControl, Enable, But8
	 GuiControl, Enable, Edi8
	 GuiControl, Enable, Chbx8

	 GuiControl, Enable, But9
	 GuiControl, Enable, Edi9
	 GuiControl, Enable, Chbx9

	 GuiControl,,Edi4, %cPath_TM1sd_exe%
	 }
  return


Changed_Editbox:
  gui, Submit, NoHide
  If (a_guicontrol = "Edi1") {

     IfExist, %Edi1%
	    GuiControl, , Chbx1, 1
     Else
	    GuiControl, , Chbx1, 0

     c01 = %Edi1%\tm1s.cfg
     IfExist, %c01%
	    {
		; retrieve the data dir and logging
		IniRead, DataDirFolder,    %c01%, TM1S, DataBaseDirectory,
        If ( InStr( DataDirFolder, ".", false ) > 0 ) OR ( InStr( DataDirFolder, "..", false ) > 0 )
        {
            base = %Edi1%
            rel = %DataDirFolder%
            DataDirFolder := % PathCombine(base, rel)
		}
        IfExist, %DataDirFolder%
            {
			GuiControl, , Chbx2, 1
            GuiControl, , Edi2, %DataDirFolder%
			}
        Else
            {
			GuiControl, , Chbx2, 0
			}

        IniRead, LoggingDirFolder,    %c01%, TM1S, LoggingDirectory,
        If ( InStr( LoggingDirFolder, ".", false ) > 0 ) OR ( InStr( LoggingDirFolder, "..", false ) > 0 )
        {
            base = %Edi1%
            rel = %LoggingDirFolder%
            LoggingDirFolder := % PathCombine(base, rel)
		}
        IfExist, %LoggingDirFolder%
            {
			GuiControl, , Chbx3, 1
			GuiControl, , Edi3, %LoggingDirFolder%
			}
        Else
            GuiControl, , Chbx3, 0

		x = %Edi1%
        StringGetPos, p1, x, \ , R1
        Stringtrimleft, vCustomer, x, (p1+1)
        GuiControl, , Edi5, %vCustomer%

        ; for the path of the shortcut, take the parent folder
		p := GetParentDir(x, 1)
        GuiControl, , Edi6, %p%

        IniRead, TM1ServerName       , %c01%, TM1S, ServerName, E
        GuiControl, , Edi7, %TM1ServerName%
	    }
     Else
     {
		If Edi1 <> C:\temporary
		{
		GuiControl, , Edi2,
		GuiControl, , Edi3,
		GuiControl, , Edi7,
		}
     }
  }
  If (a_guicontrol = "Edi2") {
     IfExist, %Edi2%
	    GuiControl, , Chbx2, 1
     Else
	    GuiControl, , Chbx2, 0
  }
  If (a_guicontrol = "Edi3") {
     IfExist, %Edi3%
	    GuiControl, , Chbx3, 1
     Else
	    GuiControl, , Chbx3, 0
  }
  If (a_guicontrol = "Edi4") {
     IfExist, %Edi4%
	    GuiControl, , Chbx4, 1
     Else
	    GuiControl, , Chbx4, 0
  }
  If (a_guicontrol = "Edi5") {
     If( Trim(Edi5) <> "" )
         {
         IfNotExist, %Edi6%\%Edi5%.lnk
	        GuiControl, , Chbx5, 1
         Else
	        GuiControl, , Chbx5, 0
		}
     Else
	    GuiControl, , Chbx5, 0
  }
  If (a_guicontrol = "Edi6") {
     IfExist, %Edi6%
	    GuiControl, , Chbx6, 1
     Else
	    GuiControl, , Chbx6, 0
  }
  If (a_guicontrol = "Edi7") {
     If( Trim(Edi6) <> "" )
	    GuiControl, , Chbx7, 1
     Else
	    GuiControl, , Chbx7, 0
  }
  If (a_guicontrol = "Edi8") {
     If( Trim(Edi8) <> "" )
	    GuiControl, , Chbx8, 1
     Else
	    GuiControl, , Chbx8, 0
  }
  If (a_guicontrol = "Edi9") {
     If( Trim(Edi9) <> "" )
	    GuiControl, , Chbx9, 1
     Else
	    GuiControl, , Chbx9, 0
  }

return

tm1_as_appGuiEscape:
  gui, destroy
  ; ExitApp
return
}

	else if(menuItem = "Convert to Expand")
    {
        Send {Esc 2}
        ; Gosub convert_to_expand

{
Clipboard := ""
Send, ^c
ClipWait, 2
StrIn := Clipboard

OpeningQuote := ""
ClosingQuote := ""
If SubStr(StrIn, 1, 1 ) != "'"
   OpeningQuote := "'"
If SubStr(StrIn, 0, 1 ) != "'"
   ClosingQuote := "'"

StrOut := RegExReplace(StrIn , "('\s*''\s*)\|\s*(\w+)\s*\|(\s*''\s*')", "''%$2%''")
StrOut := "Expand( " . OpeningQuote . RegExReplace(StrOut, "'\s*\|\s*(\w+)\s*\|\s*'", "%$1%") . ClosingQuote . " )"

Clipboard := % StrOut
Send ^v
}
        return
    }
    else if(menuItem = "Convert to | delimited")

{
Clipboard := ""
Send, ^c
ClipWait, 2
StrIn := Clipboard

StrOut := RegExReplace(StrIn, "%(.*?)%", "' | $1 | '")
Clipboard := % StrOut
Send ^v
return
    }

    If (nFound=0)
        Msgbox The selected TI menu item has an issue.

return


CountLines(Text)
{
     StringReplace, Text, Text, `n, `n, UseErrorLevel
     Return ErrorLevel + 1
}

StringBetween( String, NeedleStart, NeedleEnd )
{
    StringGetPos, pos, String, % NeedleStart
    If ( ErrorLevel )
         Return ""
    StringTrimLeft, String, String, pos + StrLen( NeedleStart )
    If ( NeedleEnd = "" )
        Return String
    StringGetPos, pos, String, % NeedleEnd
    If ( ErrorLevel )
        Return ""
    StringLeft, String, String, pos
    Return String
}

Get_Process_Name( Server, Process )
{
    if ( Process <> "___Current process" )
       return %Process%

    WinGetTitle, Window_Title, A
    if InStr( Window_Title, "Turbo Integrator:", false ) = 1
    {
        msgbox Please use Notepad++
        return
    }
    else if ( InStr( Window_Title, "Notepad++", false ) > 0 )
    {
       Arr_Caption := StrSplit(Window_Title, " - Notepad++")
	   vFile := trim(Arr_Caption[1])

	   IfInString, vFile, *
		  vFile := % Substr( vFile, 2 )
	   IfExist, %vFile%
		  Return %vFile%
    }
return
}

Get_Process_FullFileName( Server, Process )
{
    if ( Process <> "___Current process" )
       return %Process%

    WinGetTitle, Window_Title, A
    if InStr( Window_Title, "Turbo Integrator:", false ) = 1
    {
        Arr_Caption := StrSplit(Window_Title, "->")
        vProcess := trim(Arr_Caption[2])
        If( Server = "" )
             {
             Server := trim(Arr_Caption[1])
             Arr_Server := StrSplit(Server, "Turbo Integrator:")
             Server := trim(Arr_Server[2])
             }
        DataDir := GetSetting_ByServer( Server, "tm1datadirectory" )
        vFile = %DataDir%%vProcess%.pro
        IfExist, %vFile%
			return %vFile%
        Else
			return
    }
    else if ( InStr( Window_Title, "Notepad++", false ) > 0 )
    {
       DataDir := GetSetting_ByServer( Server, "tm1datadirectory" )
       If( DataDir = "" )
       {
           Arr_Caption := StrSplit(Window_Title, " - Notepad++")
           vFile := trim(Arr_Caption[1])

           ; vFile = %Window_Title%
           ; vFile := StrReplace(vFile, " - Notepad++", "")
           ; vFile := StrReplace(vFile, " [Administrator]", "")
           IfInString, vFile, *
              vFile := % Substr( vFile, 2 )
           IfExist, %vFile%
              Return %vFile%
       }
       else
	   {
		   Arr_Caption := StrSplit(Window_Title, " - Notepad++")
           vFile := trim(Arr_Caption[1])

           ; vFile = %Window_Title%
           ; vFile := StrReplace(vFile, " - Notepad++", "")
           ; vFile := StrReplace(vFile, " [Administrator]", "")
           IfInString, vFile, *
              vFile := % Substr( vFile, 2 )
           IfExist, %vFile%
           {
              SplitPath, vFile, , vDir, OutExtension, Filename
              vDir .= "\"
              If( OutExtension = "pro" )
                 If( DataDir = vDir )
                    return %Filename%
           }
       }
    }
return
}

Get_TI_Code_Tab( )
{
    Send {Esc 2}

    WinGetTitle, TM1_Window_Title, A

    ; select the contents

    WinGet, hWnd, ID, A
    SendMessage, 0x130B,,, SysTabControl327, % "ahk_id " hWnd ;TCM_GETCURSEL := 0x130B
    vSelection := ErrorLevel+1

    if (vSelection <> 4)
    {
       msgbox Select the Advanced tab (you now selected %vSelection%)
       return
    }

    WinGet, hWnd, ID, A
    SendMessage, 0x130B,,, SysTabControl321, % "ahk_id " hWnd ;TCM_GETCURSEL := 0x130B
    vSelection := ErrorLevel+1

    if (vSelection < 2)
    {
       msgbox Select Prolog, Metadata, Data or Epilog (you now selected %vSelection%)
       return
    }

    ; get the contents of the selected tab
    WinGet, hWnd, ID, A
    WinGet, vCtlList, ControlList, % "ahk_id " hWnd

    vSelection = % vSelection-1
    ControlGetText, vText, Edit%vSelection%, % "ahk_id " hWnd
    vText := StrReplace(vText, "#****Begin: Generated Statements***", "")
    vText := StrReplace(vText, "#****End: Generated Statements****", "")
    vText := Trim(vText, "`r`n")

    ;add the Advanced tab name in front
    vLabel := ""
    oArray := ["Prolog","Metadata","Data","Epilog"]
    if vSelection in 1,2,3,4
        vLabel := oArray[vSelection]

    vLabel = %vSelection% %vLabel%

    Return % vText
}
Return

Add_Text(s1, s2, m)
{
    If( m = 1 )
		erbij := s2
	Else If( m = 2 )
		erbij := "`r`n" . s2
	Else If( m = 3 )
	{
		j := StrLen(s1) + 1 + StrLen("`r`n") - InStr(s1, "`r`n", false, -1, 1)
		If( j + StrLen(s2) <= 100 )
			erbij := s2
		Else
			erbij := "`r`n" . s2
	}
    return erbij
}
return

PathCombine(abs, rel) {
    VarSetCapacity(dest, (A_IsUnicode ? 2 : 1) * 260, 1) ; MAX_PATH
    DllCall("Shlwapi.dll\PathCombine", "UInt", &dest, "UInt", &abs, "UInt", &rel)
    Return, dest
}

GetParentDir(Path,Count=1,Delimiter="\") {
  While (InStr(Path,Delimiter) <> 0 && Count <> A_Index - 1)
    Path := SubStr(Path,1,InStr(Path,Delimiter,0,0) - 1)
  Return Path
    }

GetSelectedFileFolderNames_UnderCursor()
{
	ToReturn =
	hwnd := hwnd ? hwnd : WinExist("A")
	WinGetClass class, ahk_id %hwnd%
	if (class="CabinetWClass" or class="ExploreWClass")
	{
		for window in ComObjCreate("Shell.Application").Windows
			if (window.hwnd==hwnd)
	        {
               sel := window.Document.SelectedItems
	           for item in sel
	           {
                  v := % item.path
			      ToReturn .= v "`n"
               }
            }
    }
	else
    {
	   files:={}
	   sel:=""
	   if (class="Progman" or class="WorkerW")
          FolderPath = %A_Desktop%
       else if WinActive("ahk_class CabinetWClass")
		  ; http://ahkscript.org/boards/viewtopic.php?p=28751#p28751
		  for window in ComObjCreate("Shell.Application").Windows
		      try FolderPath := window.Document.Folder.Self.Path
       ControlGet, win, HWND,, SysListView321, A
       ControlGet, sel, List, Selected Col1,,ahk_id %win%
       Loop, parse, sel, `n, `r
          files.push(A_LoopField)
       for each, file in files
       {
          FileGetAttrib, FileAttrib, %FolderPath%\%file%
          If InStr(FileAttrib, "D") ; a DIRECTORY
          {
             Loop, %FolderPath%\%file%, 2,0	; only folders
             {
                StringReplace, FileFullPath, A_LoopFileFullPath, :\\ ,:\
                ; MsgBox, DIRECTORY: "%FileFullPath%"
                v = %FileFullPath%
			    ToReturn .= v "`n"
             }
          }
          else
          {
             Loop, %FolderPath%\%file%.*
             {
                ; MsgBox, FILE: "%A_LoopFileFullPath%"
                v := % A_LoopFileFullPath
		        ToReturn .= v "`n"
             }
          }
       }
	}
    ToReturn := Trim(ToReturn,"`n")
	return ToReturn
}

GetSelectedFileFolderNames_FolderYouSee()
{
    ToReturn =
    ; WinGetClass, class, A
	hwnd := hwnd ? hwnd : WinExist("A")
	WinGetClass class, ahk_id %hwnd%
    if (class="CabinetWClass" or class="ExploreWClass")
    {
        ControlGetText, currentpath, ToolbarWindow323, ahk_class %class%
        StringTrimLeft, currentpath, currentpath, 9
        if( SubStr( currentpath, 0, 1) == "\")
           currentpath := SubStr(currentpath, 1, StrLen( currentpath ) - 1 )
        ToReturn .= currentpath "`n"
        ToReturn := Trim(ToReturn,"`n")
        Return ToReturn
    }
	else if (class="Progman" or class="WorkerW")
    {
        currentpath = %A_Desktop%
        ToReturn .= currentpath "`n"
        ToReturn := Trim(ToReturn,"`n")
        Return ToReturn
    }
    Return
}

ExpandEnvVars(sEV)
{
	VarSetCapacity(dest, 2000)
	DllCall("ExpandEnvironmentStrings", "str", sEV, "str", dest, int, 1999, "Cdecl int")
	return dest
}
return

ReplaceEnvVars(sLocation)
{
	if ( InStr( sLocation, "%HO%", false ) > 0 )
	   sLocation := StrReplace(sLocation, "%HO%", ExpandEnvVars("%HO%"))
	return sLocation
}
return

fRework_Path(s)
{
    s := StrReplace(s, "https://aexisnv-my.sharepoint.com/personal/wgielis_aexis_com/Documents/", "%HO%\")
    s := StrReplace(s, "/", "\")
    s := ReplaceEnvVars(s)
    return s
}
return

InsertEnvVars(sLocation)
{
	if ( InStr( sLocation, ExpandEnvVars("%HO%"), false ) > 0 )
	   sLocation := StrReplace(sLocation, ExpandEnvVars("%HO%"), "%HO%")
	return sLocation
}
return

; https://autohotkey.com/board/topic/60723-can-autohotkey-retrieve-file-path-of-the-selected-file/page-2
Explorer_GetSelection(hwnd="") {
	hwnd := hwnd ? hwnd : WinExist("A")
	WinGetClass class, ahk_id %hwnd%
	if (class="CabinetWClass" or class="ExploreWClass" or class="Progman")
		for window in ComObjCreate("Shell.Application").Windows
			if (window.hwnd==hwnd)
    sel := window.Document.SelectedItems
	for item in sel
	{
    ToReturn .= item.path "`n"
    Break
    }
	return Trim(ToReturn,"`n")
}
return