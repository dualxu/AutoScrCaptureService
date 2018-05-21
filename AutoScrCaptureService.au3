; Attach to the 3rd browser control embedded in another window
; Use the advanced window title syntax to use the 2nd window
; with the string 'ICQ' in the title

#include <IE.au3>
#include <ScreenCapture.au3>
#include <AutoItConstants.au3>

;打开默认浏览器-
RunWait('"' & @ComSpec & '" /c explorer.exe http://124.72.48.52:8080/mainpage', '', @SW_HIDE)
Sleep(5000)
;ToolTip("已打开目标网页", 0, 0, "Thanks", 1)
;Sleep(1000)

;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;删除Images目录下3个文件
$myFile = "D:\WorkSpace\Work2016\厦门采书柜\tools\AutoScrCaptureService\Images\ImageCapture1.jpg"
If FileExists($myFile) Then
    FileDelete($myFile)
EndIf
Sleep(2000)

$myFile = "D:\WorkSpace\Work2016\厦门采书柜\tools\AutoScrCaptureService\Images\ImageCapture2.jpg"
If FileExists($myFile) Then
    FileDelete($myFile)
EndIf
Sleep(2000)

$myFile = "D:\WorkSpace\Work2016\厦门采书柜\tools\AutoScrCaptureService\Images\ImageCapture3.jpg"
If FileExists($myFile) Then
    FileDelete($myFile)
EndIf
Sleep(2000)



;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;开始截图

;++++借阅统计数据
;业务统计
MouseMove(410,63+100,3);
Sleep(1000);
;图书借阅
MouseMove(410,176+100,3);
MouseClick($MOUSE_CLICK_LEFT);
MsgBox(0,"提示","赶紧选择日期段，给你20秒钟！",1);
Sleep(20000);
;选择日期
MsgBox(0,"提示","选择完了吗，来不及了！",1);
;采书柜
MouseMove(940,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;预约柜
MouseMove(1123,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;统计
MouseMove(1251,135+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(3000);
;截屏
MouseMove(344,100+100,3);
Sleep(1000);
MouseMove(1354,624+100,3);
Sleep(1000);
_ScreenCapture_Capture("D:\WorkSpace\Work2016\厦门采书柜\tools\AutoScrCaptureService\Images\ImageCapture1.jpg",344,100+100,1354,624+100);

;++++上架和下架统计数据
;业务统计
MouseMove(410,63+100,3);
Sleep(1000);
;图书上下架
MouseMove(410,156+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;采书柜
MouseMove(940,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;预约柜
MouseMove(1123,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;统计
MouseMove(1251,135+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(3000);
;
MouseMove(344,100+100,3);
Sleep(1000);
MouseMove(1354,624+100,3);
Sleep(1000);
_ScreenCapture_Capture("D:\WorkSpace\Work2016\厦门采书柜\tools\AutoScrCaptureService\Images\ImageCapture2.jpg",344,100+100,1354,624+100);

;++++当前在架统计数据
;业务统计
MouseMove(410,63+100,3);
Sleep(1000);
;当前在架
MouseMove(410,110+100,3);
MouseClick($MOUSE_CLICK_LEFT);
;统计
MouseMove(1117,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(3000);
;
MouseMove(344,100+100,3);
Sleep(1000);
MouseMove(1354,624+100,3);
Sleep(1000);
_ScreenCapture_Capture("D:\WorkSpace\Work2016\厦门采书柜\tools\AutoScrCaptureService\Images\ImageCapture3.jpg",344,100+100,1354,624+100);

;++++++++++++++++++++++++++++++++++++++++++++
;打开Imges文件夹，选中并拷贝3个图片文件
RunWait('"' & @ComSpec & '" /c explorer.exe D:\WorkSpace\Work2016\厦门采书柜\tools\AutoScrCaptureService\Images', '', @SW_HIDE)
Sleep(1000);
$handle = WinGetHandle("Images","");
WinActivate($handle);
Sleep(1000);
$pos = WinGetPos($handle);
;MsgBox(0,"pos[0]",$pos[0],3);
;Sleep(1000);
;MsgBox(0,"pos[1]",$pos[1],3);
;Sleep(1000);
MouseMove($pos[0]+657,$pos[1]+556,3);
MouseClick("");
Sleep(1000);
;ControlSend($handle,"","DirectUIHWND3","^a");
Send("^a")
Sleep(1000);
;ControlSend($handle,"","DirectUIHWND3","^c");
Send("^c")
Sleep(1000);

;++++++++++++++++++++++++++++++++++++++++++++
;微信上搜索并点开相应微信群，然后拷贝，发送
$handle = WinGetHandle("[CLASS:WeChatMainWndForPC]", "");
WinActivate($handle);
Sleep(1000);
$pos = WinGetPos($handle);
MouseMove($pos[0]+100,$pos[1]+30,3);
MouseClick("");
Sleep(1000)
;ControlSend("微信","","","厦门市图采书柜");
Send("厦门市图采书柜");
Sleep(1000);
MouseMove($pos[0]+160,$pos[1]+100,3);
MouseClick("");
Sleep(1000);
;
MouseMove($pos[0]+640,$pos[1]+520,3);
MouseClick("");
;ControlSend("微信","","","JJJJJJJJJJ{Enter}"); ;OKay
Sleep(1000);
;ControlSend("微信","","","^v");
Send("^v");
Sleep(1000)
;点发送
;ControlSend("微信","","","{Enter}");
Send("{Enter}");
Sleep(3000)
;发送提示
;ControlSend("微信","","","以上是厦门市图试运行采书柜和预约柜上一周借阅、上下架及当前在架图书统计。{Enter}");
Send("以上是厦门市图试运行采书柜和预约柜上一周借阅、上下架及当前在架图书统计。{Enter}");

MsgBox(0,"提示","全部完成，赞！！！",5);

