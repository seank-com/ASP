<%@ Language=VBScript %>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Generates a maze that has exactly one solution.
'
Option Explicit

Dim Maze()

Const MazeWidth = 33
Const MazeHeight = 38

Class Cell
	Public fWallUp
	Public fWallDown
	Public fWallLeft
	Public fWallRight
	Public fBorderUp
	Public fBorderDown
	Public fBorderLeft
	Public fBorderRight
	Public nBackTrackX
	Public nBackTrackY

	Private Sub Class_Initialize
		fWallUp = True
		fWallDown = True
		fWallLeft = True
		fWallRight = True
		fBorderUp = False
		fBorderDown = False
		fBorderLeft = False
		fBorderRight = False
		nBackTrackX = 0
		nBackTrackY = 0
	End Sub
End Class

Sub InitMaze
	Dim x,y
	ReDim Maze(MazeWidth, MazeHeight)

	For x = 1 to MazeWidth
		For y = 1 to MazeHeight
			set Maze(x,y) = New Cell
		Next
	Next

	For x = 1 to MazeWidth
		Maze(x,1).fBorderUp = True
		Maze(x,MazeHeight).fBorderDown = True
	Next

	For y = 1 to MazeHeight
		Maze(1, y).fBorderLeft = True
		Maze(MazeWidth, y).fBorderRight = True
	Next

	Maze(1,1).fWallUp = False
	Maze(MazeWidth,MazeHeight).fWallDown = False
End Sub

Function GetCell(fWallUp, fWallDown, fWallLeft, fWallRight)
	Dim nBlockID

	nBlockID = 0
	If fWallUp Then nBlockID = nBlockID + 8
	If fWallDown Then nBlockID = nBlockID + 4
	If fWallLeft Then nBlockID = nBlockID + 2
	If fWallRight Then nBlockID = nBlockID + 1

	Select Case nBlockID
	Case 0
	    GetCell = "<td>&nbsp;</td>"
	Case 1
	    GetCell = "<td class=""R"">&nbsp;</td>"
	Case 2
	    GetCell = "<td class=""L"">&nbsp;</td>"
	Case 3
	    GetCell = "<td class=""LR"">&nbsp;</td>"
	Case 4
	    GetCell = "<td class=""D"">&nbsp;</td>"
	Case 5
	    GetCell = "<td class=""DR"">&nbsp;</td>"
	Case 6
	    GetCell = "<td class=""DL"">&nbsp;</td>"
	Case 7
	    GetCell = "<td class=""DLR"">&nbsp;</td>"
	Case 8
	    GetCell = "<td class=""U"">&nbsp;</td>"
	Case 9
	    GetCell = "<td class=""UR"">&nbsp;</td>"
	Case 10
	    GetCell = "<td class=""UL"">&nbsp;</td>"
	Case 11
	    GetCell = "<td class=""ULR"">&nbsp;</td>"
	Case 12
	    GetCell = "<td class=""UD"">&nbsp;</td>"
	Case 13
	    GetCell = "<td class=""UDR"">&nbsp;</td>"
	Case 14
	    GetCell = "<td class=""UDL"">&nbsp;</td>"
	Case 15
	    GetCell = "<td class=""UDLR"">&nbsp;</td>"
	End Select
End Function

Sub GoUp(ByRef x, ByRef Y)
	Maze(x,Y - 1).nBackTrackX = x
	Maze(x,Y - 1).nBackTrackY = Y
	Maze(x,Y).fWallUp = False
	Maze(x,Y - 1).fWallDown = False
	Y = Y - 1
End Sub

Sub GoDown(ByRef x, ByRef Y)
	Maze(x,y+1).nBackTrackX = x
	Maze(x,y+1).nBackTrackY = y
	Maze(x,y).fWallDown = False
	Maze(x,y+1).fWallUp = False
	y = y + 1
End Sub

Sub GoLeft(ByRef x, ByRef Y)
	Maze(x-1,y).nBackTrackX = x
	Maze(x-1,y).nBackTrackY = y
	Maze(x,y).fWallLeft = False
	Maze(x-1,y).fWallRight = False
	x = x - 1
End Sub

Sub GoRight(ByRef x, ByRef Y)
	Maze(x+1,y).nBackTrackX = x
	Maze(x+1,y).nBackTrackY = y
	Maze(x,y).fWallRight = False
	Maze(x+1,y).fWallLeft = False
	x = x + 1
End Sub

Sub RandomWalk(ByRef x, ByRef y)
	Dim Dir

	Do While True
		Dir = Rnd * 1000 mod 4

		Select Case Dir
		Case 1 'Go Up
			If Maze(x,y).fBorderUp Then Exit Sub
			If Maze(x,y-1).nBackTrackX <> 0 Then Exit Sub
			GoUp x, y
		Case 2 'Go Down
			If Maze(x,y).fBorderDown Then Exit Sub
			If Maze(x,y+1).nBackTrackX <> 0 Then Exit Sub
			GoDown x, y
		Case 3 'Go Left
			If Maze(x,y).fBorderLeft Then Exit Sub
			If Maze(x-1,y).nBackTrackX <> 0 Then Exit Sub
			GoLeft x, y
		Case Else 'Go Right
			If Maze(x,y).fBorderRight Then Exit Sub
			If Maze(x+1,y).nBackTrackX <> 0 Then Exit Sub
			GoRight x, y
		End Select
	Loop
End Sub

Sub BackTrack(ByRef x, ByRef y)
	Dim BackX, BackY
	Do While Maze(x,y).nBackTrackX <> -1
		If Not Maze(x,y).fBorderUp Then
			If Maze(x,y-1).nBackTrackX = 0 Then
				GoUp x, y
				Exit Sub
			End If
		End If

		If  Not Maze(x,y).fBorderDown Then
			If Maze(x,y+1).nBackTrackX = 0 Then
				GoDown x, y
				Exit Sub
			End If
		End If

		If  Not Maze(x,y).fBorderLeft Then
			If Maze(x-1,y).nBackTrackX = 0 Then
				GoLeft x, y
				Exit Sub
			End If
		End If

		If  Not Maze(x,y).fBorderRight Then
			If Maze(x+1,y).nBackTrackX = 0 Then
				GoRight x, y
				Exit Sub
			End If
		End If

		BackX = Maze(x,y).nBackTrackX
		BackY = Maze(x,y).nBackTrackY
		x = BackX
		y = BackY
	Loop
End Sub

Sub GenerateMaze
	Dim x, y

	Randomize Timer
	x = MazeWidth \ 2
	y = MazeHeight \ 2
	Maze(x,y).nBackTrackX = -1

	RandomWalk x, y
	Do While True
		BackTrack x, y
		If Maze(x,y).nBackTrackX = -1 Then Exit Sub
		RandomWalk x, y
	Loop
End Sub

Sub RenderMaze
	Dim x, y

  	Response.Write "<table border=""0"" width=""100%"" height=""100%"" cellpadding=""0"" cellspacing=""0"">"

	for y = 1 to MazeHeight
    	Response.Write "<tr>"
	    for x = 1 to MazeWidth
	        Response.Write GetCell(Maze(x,y).fWallUp,Maze(x,y).fWallDown,Maze(x,y).fWallLeft,Maze(x,y).fWallRight)
    	next
    	Response.Write "</tr>"
	next

   	Response.Write "</table>"
End Sub

InitMaze
GenerateMaze

%>
<html>
<head>
</head>
<title>Random Maze</title>
<style>
   BODY {  background-color: white; color: black;  }
   TD { border:3px solid white }
   .D { border-bottom-color: black; }
   .DL { border-bottom-color: black; border-left-color: black; }
   .DLR { border-bottom-color: black; border-left-color: black; border-right-color: black; }
   .DR { border-bottom-color: black; border-right-color: black; }
   .L { border-left-color: black; }
   .LR { border-left-color: black; border-right-color: black; }
   .R { border-right-color: black; }
   .U { border-top-color: black; }
   .UD { border-top-color: black; border-bottom-color: black; }
   .UDL { border-top-color: black; border-bottom-color: black; border-left-color: black; }
   .UDLR { border-top-color: black; border-bottom-color: black; border-left-color: black; border-right-color: black; }
   .UDR { border-top-color: black; border-bottom-color: black; border-right-color: black; }
   .UL { border-top-color: black; border-left-color: black; }
   .ULR { border-top-color: black; border-left-color: black; border-right-color: black; }
   .UR { border-top-color: black; border-right-color: black; }
</style>
</head>
<body>
<%

RenderMaze

%>
</body>
</html>
