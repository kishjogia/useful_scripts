'/********************************************************************************/
'/*                                                                              */
'/*           _____________________________________________________              */
'/*           |        A D D   P R I N T E R   S C R I P T        |              */
'/*           -----------------------------------------------------              */
'/*                                                                              */
'/*  Source file name      :-      add_Printer.vbs                               */
'/*                                                                              */
'/*  Initial creation date :-      22-MAY-2013                                   */
'/*                                                                              */
'/*  Original Developers   :-      Kish Jogia                                    */
'/*                                                                              */
'/********************************************************************************/
'/*                                                                              */
'/*  Description :-                                                              */
'/*                                                                              */
'/*  Add a new Printer when you logon to a computer.                             */
'/*                                                                              */
'/********************************************************************************/

Option Explicit
On Error Resume Next

Dim objNetwork, strUNCPrinter, strLPTPrinter

strUNCPrinter = "printer"               '/* Change the name of the printer in quotes
Set objNetwork = CreateObject("WScript.Network") 
objNetwork.AddWindowsPrinterConnection strUNCPrinter

WScript.Quit

'/*  End of Windows logon script.                                                */