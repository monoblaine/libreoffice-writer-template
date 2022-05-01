sub ConvertToSmallerSourceText
    _ConvertToSmallerCharacterStyle("Source Text")
end sub

sub ConvertToSmallerHighlightedSourceText
    _ConvertToSmallerCharacterStyle("Highlighted Source Text")
end sub

sub _ConvertToSmallerCharacterStyle (charStyleName as String)
    text = ThisComponent.CurrentSelection.getByIndex(0)
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
    text = ThisComponent.CurrentSelection.getByIndex(0)
    text.CharHeight = text.CharHeight * factor
end sub

sub EndOfWord
    document = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dim args(0) as new com.sun.star.beans.PropertyValue

    EndOfWordExtend

    dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args())
    dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args())
end sub

sub EndOfWordExtend
    initialSelection = ThisComponent.CurrentSelection.getByIndex(0).String
    initialSelectionLen = Len(initialSelection)

    document = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dim args(0) as new com.sun.star.beans.PropertyValue

    dispatcher.executeDispatch(document, ".uno:WordRightSel", "", 0, args())

    newSelection = ThisComponent.CurrentSelection.getByIndex(0).String
    newSelectionLen = Len(newSelection)

    textSearch = CreateUnoService("com.sun.star.util.TextSearch")
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
