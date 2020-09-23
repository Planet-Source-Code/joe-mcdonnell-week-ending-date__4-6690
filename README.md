<div align="center">

## Week Ending Date


</div>

### Description

Returns the date of the week-ending day based on the input date. Useful for due-by dates, weekly accounts etc.
 
### More Info
 
A date, or use date() for the current date.

The number of the week-ending day (Sun=1 ... Sat=7).

Example: myvalue=<%=weekendday(date(),1)%>

A date.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joe McDonnell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joe-mcdonnell.md)
**Level**          |Advanced
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Math](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math__4-12.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joe-mcdonnell-week-ending-date__4-6690/archive/master.zip)





### Source Code

```
This function accepts the input date and the weekending day (Sun=1.....Sat=7).
If the weekday of the input date is the same as the weekending day it returns the input date
Function weekendday(inputdate, day)
If Weekday(inputdate) > day Then
inputdate = inputdate + 7
End If
weekendday = inputdate + (day - Weekday(inputdate))
End Function
```

