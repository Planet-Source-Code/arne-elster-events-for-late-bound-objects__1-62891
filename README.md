<div align="center">

## events for late\-bound objects


</div>

### Description

VB has good event handling functions.

The great WithEvents and VBControlExtender.

The bad thing about them is, WithEvents can only be used for early binding, and the VBControlExtender is only for controls.

So what to do in case you have an object like "InternetExplorer.Application" and want to catch its "Stop" event?

Right, you guessed it: grapple with some nasty COM interfaces :)

In this last part of my low level COM project I want to show you how to use IConnectionPoint to recieve events from objects.

Oh, and you may see some equalities between this submission and the event collection by Edanmo.

Edanmo has some functions in there I really was frightened to code. ;)
 
### More Info
 
don't hit that stop button when you're inside the event sink!


<span>             |<span>
---                |---
**Submitted On**   |2005-10-14 16:10:18
**By**             |[Arne Elster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/arne-elster.md)
**Level**          |Advanced
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[events\_for19404010142005\.zip](https://github.com/Planet-Source-Code/arne-elster-events-for-late-bound-objects__1-62891/archive/master.zip)








