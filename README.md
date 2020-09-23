<div align="center">

## FUNCTION: Return Astrology Sign By Birthday


</div>

### Description

Function: Returns someones astrology sign using their birthday in m/dd/yyyy or m/dd/yy format ie. 8/27/1983 -- > Virgo
 
### More Info
 
A date value

8/27/1983

8/27/83


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[snowboardr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/snowboardr.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/snowboardr-function-return-astrology-sign-by-birthday__4-8546/archive/master.zip)





### Source Code

```
<% Option Explicit
'# By Jason H
'# Improvements, comments, or if you find it useful
'# let me know jason@vzio.com
'#
'# September 15, 2003
'#
'# Working example below
'# Response.Write(UserSign("8/27/2003"))
Function UserSign(sUserDate)
	Dim sMonth
	Dim sDay
	'# Remove year
	sUserDate = Left(sUserDate,InstrRev(sUserDate,"/")-1)
	sMonth = CINT(Left(sUserDate,(Instr(sUserDate,"/")-1)))
	sDay = CINT(Replace(sUserDate, sMonth & "/", ""))
	 Select CASE sMonth
				Case 1
					If sDay => 20 then
							UserSign = "Aquarius"
							 		ElseIf sDay <= 19 then
							UserSign = "Capricorn"
					End If
				Case 2
					If sDay <= 18 then
							UserSign = "Aquarius"
								ElseIf sDay => 19 then
							UserSign = "Pisces"
					End If
				Case 3
					If sDay <= 20 then
							UserSign = "Pisces"
								ElseIf sDay => 21 then
							UserSign = "Aries"
					End If
				Case 4
					If sDay <= 19 then
							UserSign = "Aries"
								ElseIf sDay => 20 then
							UserSign = "Taurus"
					End If
				Case 5
					If sDay <= 20 then
							UserSign = "Taurus"
								ElseIf sDay => 21 then
							UserSign = "Gemini"
					End If
				Case 6
					If sDay <= 21 then
							UserSign = "Gemini"
								ElseIf sDay => 22 then
							UserSign = "Cancer"
					End If
				Case 7
					If sDay <= 22 then
							UserSign = "Cancer"
								ElseIf sDay => 23 then
							UserSign = "Leo"
					End If
				Case 8
					If sDay <= 22 then
							UserSign = "Leo"
								ElseIf sDay => 23 then
							UserSign = "Virgo"
					End If
				Case 9
					If sDay <= 22 then
							UserSign = "Virgo"
								ElseIf sDay => 23 then
							UserSign = "Libra"
					End If
				Case 10
					If sDay <= 22 then
							UserSign = "Libra"
								ElseIf sDay => 23 then
							UserSign = "Scorpio"
					End If
				Case 11
					If sDay <= 21 then
							UserSign = "Scorpio"
								ElseIf sDay => 22 then
							UserSign = "Sagittarius"
					End If
				Case 12
					If sDay <= 21 then
							UserSign = "Sagittarius"
								ElseIf sDay => 22 then
							UserSign = "Capricorn"
					End If
	End Select
End Function
Response.Write(UserSign("8/27/2003"))
%>
```

