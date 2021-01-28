; ####################################################
; #	
; # Author: Wim Gielis
; # Contact: wim.gielis@gmail.com
; # URL: https://github.com/wimgielis?tab=repositories
; # Website: https://www.wimgielis.com
; # Date: January 28, 2021
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

	toolName:= "Update PRO line numbers"

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
		; # - open up any (valid) PRO file in Notepad++
		; # - launch the menu and the gui
		; # - list all variables of a TI process
		; # - allow to generate the TI code for selected variables
		; # - it will be pasted at the cursor
		; # - several options are built-in
		; # 
		; # To be added later:
		; # - Order of the selected variables could be respected and AsciiOutput in that order
		; # - A tab control on the gui containing a 'live' preview of what the code would create (with manual adjustment)
		; # - Use the Windows registry to save to and load from for choices in the gui (not needed to redo everything)
		; # - No REST API used and needed for now, I parse a text file and that's it
		; # - Notepad++ is needed
		; ####################################################
        

		vProcess := Get_Process_Name( "", "___Current process" )

		Gui, Destroy

		Gui, Add, Text, vFullFilename Disabled, %vProcess%
		
		Gui, Add, Button, gAO_SelectALL x15, Select all
		Gui, Add, Button, gAO_ReloadVars x+30, Reload variables
		Gui, Add, ListView, x10 w400 h500, Variable Type|Variable Name|Data Type
		Gui, Add, ListBox, 8 x+10 w150 r9 Choose1 vShow_Which_Variables gShow_Which_Variables, All|Parameters|Data source variables|Data source new variables|Custom variables|   in the Prolog tab|   in the Metadata tab|   in the Data tab|   in the Epilog tab
		Gui, Add, ListBox, AltSubmit ReadOnly y+30 w150 r3 vSwitch_Text_Number gSwitch_Text_Number, Text <--> Number|--> Text|--> Number
		Gui, Add, Button, gAO_OK y+40, Generate output
		
		Gui, Add, ListBox, x10 w100 Choose1 vMyListBox_OutputFunction, AsciiOutput|TextOutput
		Gui, Add, ListBox, AltSubmit x+20 w100 Choose1 vMyListBox_OutputFolder, DataDir|LoggingDir
		Gui, Add, edit, vEdi1 x+20 w160 r1, test.txt
		Gui, Add, CheckBox, y+10 vVariable_Filename, Variable
		
        Gui, Add, CheckBox, x10 y+15 vAdd_Output_Settings, Add output settings
        Gui, Add, CheckBox, x10 vNumbers_NumberFormat, Numbers get a numberformat

        Gui, Add, CheckBox, x10 vNames_of_variables, Insert names of variables
        Gui, Add, CheckBox, x10 vHeader, Add code for a header
        Gui, Add, CheckBox, x10 vOutput_in_Selected_Order, Output variables in selected order

		Gui, Add, ListBox, AltSubmit Choose1 x10 y+15 w150 r3 vLine_Splitter, No split|Multiline output|Split long lines

		Gui, Add, Edit, x250 y650 w50 Center
		Gui, Add, UpDown, Left vUpDown_IF Range0-10, 0
		Gui, Add, Text, x+5 yp+3, Add this many filters

		Gui, Add, Edit, x250 y685 w50 Center
		Gui, Add, UpDown, Left vUpDown_Repeat Range0-20, 1
		Gui, Add, Text, x+5 yp+3, Repeat output this many times

        Gosub Read_in_AO_vars

		LV_ModifyCol()  ; Auto-size each column to fit its contents
		Gui, Show, w600 h850, Output variables to a text file                                                         (c) 2021 - Wim Gielis

        return
		
		Show_Which_Variables:
		gosub Read_in_AO_vars
		return

		Read_in_AO_vars:
		Gui, Submit, NoHide
		LV_Delete()

		guiControlGet, sFullFilename,, FullFilename
		FileRead, FileContents, %sFullFilename%


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
				
					If ( A_index < Nr_of_vars_without_builtin + 2 )
					{
						If ( my_param_type = 2 )
							LV_Add("", "Data source var", my_param_name, "Text")
						Else
							LV_Add("", "Data source var", my_param_name, "Number")
					}
					Else
					{
						If ( my_param_type = 2 )
							LV_Add("", "View var", my_param_name, "Text")
						Else
							LV_Add("", "View var", my_param_name, "Number")
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
			Text_of_vars_1234 := StringBetween( FileContents, "572,", "576," )
			Text_of_vars_1    := StringBetween( FileContents, "572,", "573," )
			Text_of_vars_2    := StringBetween( FileContents, "573,", "574," )
			Text_of_vars_3    := StringBetween( FileContents, "574,", "575," )
			Text_of_vars_4    := StringBetween( FileContents, "575,", "576," )

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
			; While Pos := RegExMatch(Text_of_vars, "m)^\s*?\w+?\s*?=", m, Pos+StrLen(m))
			; While Pos := RegExMatch(Text_of_vars, "\v\s*?\w+?\s*?=", m, Pos+StrLen(m)) ; somewhat more whitespace in output (see my AHK forum thread)
			   {
			   m := StrReplace(m, "=", "")
			   m := StrReplace(m, A_Space, "")
			   m := StrReplace(m, "`r`n", "")
			   allMatches .= (!allMatches ? "" : "`n") m
			   }

			Loop, parse, allMatches, `n
			If not RegExMatch(out, "\b" A_LoopField "\b") {
				out .= out ? "`n" : ""
				out .= A_LoopField
			}

			VarArray := StrSplit(out, "`n")
			Loop % VarArray.MaxIndex()
			{
				my_var_name := VarArray[A_Index]
				StringLower, vmy_var_name, my_var_name

				if( vmy_var_name = "datasourceasciidelimiter"
				 Or vmy_var_name = "datasourceasciiquotecharacter"
				 Or vmy_var_name = "datasourceasciidecimalseparator"
				 Or vmy_var_name = "datasourceasciithousandseparator"
				 Or vmy_var_name = "datasourceasciiheaderrecords"
				 Or vmy_var_name = "vmask"
				 Or vmy_var_name = "vdec"
				 Or vmy_var_name = "v1000" )
					Continue

				vFirst_Letter := Substr( my_var_name, 1, 1 )
				StringLower, vFirst_Letter, vFirst_Letter
				if( vFirst_Letter = "n" )
					LV_Add("", "Custom variable", my_var_name, "Number")
				else if( vFirst_Letter = "s" )
					LV_Add("", "Custom variable", my_var_name, "Text")
				else
					LV_Add("", "Custom variable", my_var_name, "Text/Number")
			}
		}
		return

		AO_SelectALL:
		Gui, Submit, NoHide
		LV_Modify(0, "Select")
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
			If MyListBox_OutputFolder = 2
				sOutput .= "GetProcessErrorFileDirectory | "
			sOutput .= "'" . Edi1 . "';" . e
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
			sOutput .= MyListBox_OutputFunction . "( "
			If ( Variable_Filename = True )
				sOutput .= "vFile"
		    Else
			{
				If MyListBox_OutputFolder = 2
					sOutput .= " GetProcessErrorFileDirectory | "
				sOutput .= "'" . Edi1 . "'"
			}

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

			sOutput .= e . MyListBox_OutputFunction . "( "

			If ( Variable_Filename = True )
				sOutput .= "vFile"
		    Else
			{
				If MyListBox_OutputFolder = 2
					sOutput .= " GetProcessErrorFileDirectory | "
				sOutput .= "'" . Edi1 . "'"
			}

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
						If ( Numbers_NumberFormat = False Or vVar_Name = "value_is_string" )
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
						If ( Numbers_NumberFormat = False Or vVar_Name = "value_is_string" )
							sOutput .= Add_Text(sOutput, ", '" . Var_Name . ": ', " . "NumberToString( " . Var_Name . " )", m)
						Else
							sOutput .= Add_Text(sOutput, ", '" . Var_Name . ": ', " . "NumberToStringEx( " . Var_Name . ", vMask, vDec, v1000 )", m)
						}
					Else
						sOutput .= Add_Text(sOutput, ", '" . Var_Name . ": ', " . Var_Name, m)
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