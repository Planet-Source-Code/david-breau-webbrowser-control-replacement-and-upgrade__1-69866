<div align="center">

## Webbrowser control replacement and upgrade


</div>

### Description

Are you sick of the webbrowser control and all its bugs and problems? Especially since the change from IE6 to IE7. Seems like microsoft goes out of its way to make programming difficult. Well this control called the IE controler is a solution that I came up with that works and works great!! With this control you can create a hidden or visible instance of IE and set it (the visible one) as a child of any control on your form like a picturebox. With this control you have easy access to all of the webbrowsers built in functions and properties as well as the underlying document. With one line of code you can have access and control to all the documents elements such as its links, images, forms, tables, table cells, etc. The possibilites will excite you.

Here are some of the routines and what they can do for you.

Functions/subs

LinkByText: returns a link in the document

whose text matches desired value

Go:     makes the browser go back,

forward, home, or search

return_links:     returns, in event, all the anchor objects in the document

return_tables:    returns, in event, all the table objects in the document

return_forms:     returns, in event, all the form objects in the document

return_images:    returns, in event, all the image objects in the document

return_tablerows:   returns, in event, all the tr objects in the document

return_tabledowns:  returns, in event, all the td objects in the document

CreateHiddenDocument: creates a hidden webpage from the url passed to it, from which you

can extract and manipulate all of the objects and elements on it.

WriteHtmlToDoc:    allows you to insert your own html within the browser

createIE:       creates a visible instance of ie and navigates to the url you

specify. You have the option of removing its titlebar, make it

unresizable, make it the child of any control on your form which

has the visible effect of embedding the ie browser seemlessly into

your application.

bShowContextMenu:   Allow or disallow the right click context menu

Events

Event IEDocReady(odoc As HTMLDocument)

Event IEdocState(state As String)

Event HiddenDocReady(odoc As HTMLDocument, surl As String)

Event HiddenDocState(state As String)

Event HiddenDocTimeout(lelapsed As Long)

Event IEDownloadStart()

Event IEDownloadDone()

Event returnedLinks(olink As HTMLAnchorElement, cnt As Integer)

Event returnedImages(oimage As HTMLImg, cnt As Integer)

Event returnedForms(oform As HTMLFormElement, cnt As Integer)

Event returnedTables(otable As HTMLTable, cnt As Integer)

Event returnedTableRows(oTr As HTMLTableRow, cnt As Integer)

Event returnedTableDowns(oTd As HTMLTableCell, cnt As Integer)

Event contextMenu()

Event mousedown(ibutton As Integer)

Event closing()

Event IEcreated()

Event processingDone(sFunctionName As String)

Event Error(sProcName As String, iErrNum As Long, sErrDescr As String)

Event NewWindow()
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2008-01-03 10:02:24
**By**             |[David Breau](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-breau.md)
**Level**          |Intermediate
**User Rating**    |4.7 (61 globes from 13 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Webbrowser209611132008\.zip](https://github.com/Planet-Source-Code/david-breau-webbrowser-control-replacement-and-upgrade__1-69866/archive/master.zip)








