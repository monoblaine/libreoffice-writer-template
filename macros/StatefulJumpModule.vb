global currentController as Object
global mouseClickHandler as Object
global anchor

sub RegisterMouseClickHandler
    anchor = Nothing

    currentController = ThisComponent.CurrentController
    mouseClickHandler = createUnoListener("MouseClickListener_", "com.sun.star.awt.XMouseClickHandler")
    currentController.addMouseClickHandler(mouseClickHandler)
end sub

sub UnregisterMouseClickHandler
    on error resume next
    currentController.removeMouseClickHandler(mouseClickHandler)
end sub

sub JumpToLink
    on error resume next
    vc = ThisComponent.CurrentController.ViewCursor

    if IsEmpty(vc.TextField) = false then
        anchor = vc.TextField.Anchor
        sourceName = vc.TextField.SourceName

        if InStr(sourceName, "__RefHeading___Toc") then
            _JumpToBookmark(sourceName)
        else
            _JumpToReference(sourceName)
        end if
    elseif vc.HyperLinkURL <> "" and IsEmpty(vc.TextParagraph) = false then
        anchor = vc.TextParagraph.Anchor
        _JumpToHyperLink
    end if
end sub

sub ReturnFromJump
    on error resume next

    if anchor is Nothing then
        exit sub
    end if

    ThisComponent.CurrentController.select(anchor)
    ThisComponent.CurrentController.ViewCursor.collapseToStart
    anchor = Nothing
end sub

sub _JumpToBookmark (sourceName)
    _JumpToAnchor(ThisComponent.Bookmarks.getByName(sourceName).getAnchor())
end sub

sub _JumpToReference (sourceName)
    _JumpToAnchor(ThisComponent.ReferenceMarks.getByName(sourceName).getAnchor())
end sub

sub _JumpToAnchor (targetAnchor)
    ThisComponent.CurrentController.select(targetAnchor)
    ThisComponent.CurrentController.ViewCursor.collapseToStart
end sub

sub _JumpToHyperLink
    document = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dim args(0) as new com.sun.star.beans.PropertyValue
    dispatcher.executeDispatch(document, ".uno:OpenHyperlinkOnCursor", "", 0, args())
end sub

function MouseClickListener_mousePressed (e) as Boolean
end function

function MouseClickListener_mouseReleased (e) as Boolean
    on error resume next
    vc = ThisComponent.CurrentController.ViewCursor

    if IsEmpty(vc.TextField) = false then
        anchor = vc.TextField.Anchor
    elseif vc.HyperLinkURL <> "" and IsEmpty(vc.TextParagraph) = false then
        anchor = vc.TextParagraph.Anchor
    end if
end function

sub MouseClickListener_disposing (e)
end sub
