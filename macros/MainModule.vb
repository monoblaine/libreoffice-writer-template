sub ConvertToSmallerSourceText
    _ConvertToSmallerCharacterStyle("Source Text")
end sub

sub ConvertToSmallerHighlightedSourceText
    _ConvertToSmallerCharacterStyle("Highlighted Source Text")
end sub

sub _ConvertToSmallerCharacterStyle (charStyleName as String)
    text = ThisComponent.CurrentController.Selection.getByIndex(0)
    text.CharStyleName = charStyleName
    text.CharHeight = text.CharHeight * .9
end sub

sub IncreaseFontSize
    _ChangeFontSize(1.1)
end sub

sub DecreaseFontSize
    _ChangeFontSize(0.9)
end sub

sub _ChangeFontSize (factor as Single)
    text = ThisComponent.CurrentController.Selection.getByIndex(0)
    text.CharHeight = text.CharHeight * factor
end sub

sub SetColorDanger
    _SetColor(&hFF0000)
end sub

sub SetColorSuccess
    _SetColor(&h127622)
end sub

sub SetColorWarning
    _SetColor(&hEA7500)
end sub

sub RemoveColor
    _SetColor(-1)
end sub

sub _SetColor(color as Long)
    ThisComponent.CurrentController.Selection.getByIndex(0).CharColor = color
end sub

sub EndOfWord
    on error resume next

    document = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dim args(0) as new com.sun.star.beans.PropertyValue

    EndOfWordExtend

    dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args())
    dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args())
end sub

sub EndOfWordExtend
    on error resume next

    initialSelection = ThisComponent.CurrentController.Selection.getByIndex(0).String
    initialSelectionLen = Len(initialSelection)

    document = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dim args(0) as new com.sun.star.beans.PropertyValue

    dispatcher.executeDispatch(document, ".uno:WordRightSel", "", 0, args())

    newSelection = ThisComponent.CurrentController.Selection.getByIndex(0).String
    newSelectionLen = Len(newSelection)

    textSearch = createUnoService("com.sun.star.util.TextSearch")
    textSearchOptions = CreateUnoStruct("com.sun.star.util.SearchOptions")
    textSearchOptions.algorithmType = com.sun.star.util.SearchAlgorithms.REGEXP
    textSearchOptions.searchString = "([\s]+)$"
    textSearch.setOptions(textSearchOptions)
    searchResult = textSearch.searchForward(newSelection, 0, newSelectionLen)

    if (searchResult.subRegExpressions = 0) then
        exit sub
    end if

    match = mid(newSelection, searchResult.startOffset(0) + 1, searchResult.endOffset(0) - searchResult.startOffset(0))
    numOfWhitespace = Len(match)

    if (numOfWhitespace > 0 and initialSelectionLen = (newSelectionLen - numOfWhitespace)) then
        EndOfWordExtend
        exit sub
    end if

    for i = 1 to numOfWhitespace
        dispatcher.executeDispatch(document, ".uno:CharLeftSel", "", 0, args())
    next i
end sub

' Credits: https://forum.openoffice.org/en/forum/viewtopic.php?p=355072&sid=5b308cb7f91a95563cba0257b85389d0#p355072

sub FormatCrossReferences
    dim oParEnum           'Paragraph enumerator
    dim oSecEnum           'text section enumerator
    dim oPar               'Current paragraph
    dim oParSection        'Current Section

    oParEnum = thisComponent.Text.createEnumeration()

    do while oParEnum.hasMoreElements()
        oPar = oParEnum.nextElement()

        if oPar.supportsService("com.sun.star.text.Paragraph") then
            oSecEnum = oPar.createEnumeration()

            do while oSecEnum.hasMoreElements()
                oParSection = oSecEnum.nextElement()

                if (oParSection.TextPortionType = "TextField") then
                    if (oParSection.TextField.SupportedServiceNames(0) = "com.sun.star.text.TextField.GetReference") then
                        ' set style --> warning: must exist, otherwise exception (illegal argument ...)
                        oParSection.CharStyleName = "Internet Link"
                    end if
                end if
            loop
        end if
    loop
end sub

sub RefineTextButton
   dim mCurSelection

   oDoc = thisComponent ' global defined
   oVC = oDoc.currentController.viewCursor

   mCurSelection = oDoc.currentController.Selection ' store current view cursor

   FormatCrossReferences

   oDoc.currentController.select(mCurSelection) ' restore initial cursor position

   ' MsgBox "Finished.", 0, "RefineText"
end sub

' Credits: https://forum.openoffice.org/en/forum/viewtopic.php?p=241033&sid=2be543144fe107264b87252536b76698#p241033
sub DeleteCurrentParagraph
    oDoc = ThisComponent
    oVC = oDoc.CurrentController.getViewCursor
    oTC = oDoc.Text.createTextCursorByRange(oVC)
    oTC.gotoStartOfParagraph(false)
    oTC.gotoEndOfParagraph(true)
    oTC.String = ""
    oTC.goLeft(1, true)
    oTC.String = ""

    dim document   as Object
    dim dispatcher as Object

    document   = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    dim args1(1) as new com.sun.star.beans.PropertyValue

    args1(0).Name = "Count"
    args1(0).Value = 1
    args1(1).Name = "Select"
    args1(1).Value = false

    dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args1())
end sub

sub OnCharLeft
    oVC = thisComponent.getCurrentController.getViewCursor

    if Len(oVC.String) then
        oVC.collapseToStart
    else
        oVC.goLeft(1, false)
    end if
end sub

sub OnCharRight
    oVC = thisComponent.getCurrentController.getViewCursor

    if Len(oVC.String) then
        oVC.collapseToEnd
    else
        oVC.goRight(1, false)
    end if
end sub

' Credits: https://wiki.documentfoundation.org/Macros/Writer/006
sub UpdateIndexes()
    '''Update indexes, such as for the table of contents'''

    dim i as Integer

    with ThisComponent ' Only process Writer documents
        if .supportsService("com.sun.star.text.GenericTextDocument") then
            for i = 0 to .getDocumentIndexes().count - 1
                .getDocumentIndexes().getByIndex(i).update()
            next i
        end if
    end with ' ThisComponent
end sub

sub GoToNextElementAndUnfold
    document = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dim args(0) as new com.sun.star.beans.PropertyValue
    vc = ThisComponent.CurrentController.ViewCursor
    lastPosY = 0

    do while vc.Position.Y > lastPosY
        selection = ThisComponent.CurrentController.Selection.getByIndex(0)
        outlineContentVisible = selection.OutlineContentVisible

        if UBound(outlineContentVisible) = 0 then
            if outlineContentVisible(0).Value = False then
                dispatcher.executeDispatch(document, ".uno:ToggleOutlineContentVisibility", "", 0, args())
            end if
        end if

        lastPosY = vc.Position.Y
        dispatcher.executeDispatch(document, ".uno:ScrollToNext", "", 0, args())
    loop
end sub
