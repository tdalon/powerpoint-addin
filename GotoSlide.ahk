; GoToSlide : Open Slide of ActivePresentation matching selected text by name in SlideShowView
; Selected text can start with ppt- prefix; it will be removed

^!p:: ;Ctrl+Alt+P

; pptApp := ComObjCreate("PowerPoint.Application")
 pptApp :=ComObjActive("PowerPoint.Application")
 
; Parse current line for PPT SlideName
Click 2 ; Select full word
Send ^c ; Copy to clipboard

SlideName := clipboard
; Remove prepending ppt-
SlideName := RegExReplace(SlideName,"ppt-","")

; MsgBox % SlideName

; Convert SlideName to SlideIndex
SlideIndex := 0
Loop, % pptApp.ActivePresentation.Slides.Count
{
	if  (pptApp.ActivePresentation.Slides(a_index).Name == SlideName) {
		SlideIndex := a_index	
		break
	}
		
}
MsgBox % SlideIndex
If SlideIndex = 0 	
	Return


If pptApp.SlideShowWindows.Count = 0 
    pptApp.ActivePresentation.SlideShowSettings.Run

pptApp.SlideShowWindows(1).View.GotoSlide(SlideIndex)

; change focus to power point viewer title
IfWinExist % pptApp.ActivePresentation.Name
WinActivate