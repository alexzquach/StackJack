Attribute VB_Name = "Module1"
Option Explicit

' Declarations and utilities for using CARDS.DLL
' Actions for CdtDraw/Ext (use in the nDraw field)
Global Const C_FACES = 0
Global Const C_BACKS = 1
Global Const C_INVERT = 2

' Card Numbers (use in the nCard field)
' From 0 to 51    [Ace (club,diamond,heart,spade), Deuce, ... , King]

' Card Backs (use in the nCard field)
' CAUTION: when nCard > 53 then nDraw must be = 1 (C_BACKS)
Global Const CrossHatch = 53
Global Const Weave1 = 54
Global Const Weave2 = 55
Global Const Robot = 56
Global Const Flowers = 57
Global Const Vine1 = 58
Global Const Vvine2 = 59
Global Const Fish1 = 60
Global Const Fish2 = 61
Global Const Shells = 62
Global Const Castle = 63
Global Const Island = 64
Global Const CardHand = 65
Global Const UNUSED = 66
Global Const THE_X = 67
Global Const THE_O = 68

' Initialization - Call before anything else
' Returns the default width and height for the cards, in pixels.
Declare Function cdtInit Lib "CARDS32.DLL" (nWidth As Long, nHeight As Long) As Long

' CdtDrawExt is used to draw a card in any size
' Similar to CdtDraw except that you can specify the height and width
' of the card, as well as the location.
' nWidth  = Width of card in pixels
' nHeight = Height of card in pixels
Declare Function cdtDrawExt Lib "CARDS32.DLL" (ByVal hDC As Long, ByVal xOrg As Long, ByVal yOrg As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nCard As Long, ByVal nDraw As Long, ByVal nColor As Long) As Long
  
' CdtDraw is used to draw a card with the default size at a specified location in a
' form, picture box or other control.  It can draw any of the 52 faces and 13
' different back designs, as well as pile markers such as the X and O.  Cards can
' also be drawn in the negative image; e.g. to show selection.
' xOrg  = x origin in pixels
' yOrg  = y origin in pixels
' nCard = one of the Card Back constants or a card number 0 to 51
' nDraw = one of the Action constants
' nColor = The highlight color
Declare Function cdtDraw Lib "CARDS32.DLL" (ByVal hDC As Long, ByVal xOrg As Long, ByVal yOrg As Long, ByVal nCard As Long, ByVal nDraw As Long, ByVal nColor As Long) As Long
  
' CdtTerm should be called when the program terminates
' It is used to primarily release memory back to Windows

Declare Function cdtTerm Lib "CARDS32.DLL" () As Long

