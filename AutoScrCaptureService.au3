; Attach to the 3rd browser control embedded in another window
; Use the advanced window title syntax to use the 2nd window
; with the string 'ICQ' in the title

#include <IE.au3>
#include <ScreenCapture.au3>
#include <AutoItConstants.au3>

;��Ĭ�������-
RunWait('"' & @ComSpec & '" /c explorer.exe http://124.72.48.52:8080/mainpage', '', @SW_HIDE)
Sleep(5000)
;ToolTip("�Ѵ�Ŀ����ҳ", 0, 0, "Thanks", 1)
;Sleep(1000)

;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;ɾ��ImagesĿ¼��3���ļ�
$myFile = "D:\WorkSpace\Work2016\���Ų����\tools\AutoScrCaptureService\Images\ImageCapture1.jpg"
If FileExists($myFile) Then
    FileDelete($myFile)
EndIf
Sleep(2000)

$myFile = "D:\WorkSpace\Work2016\���Ų����\tools\AutoScrCaptureService\Images\ImageCapture2.jpg"
If FileExists($myFile) Then
    FileDelete($myFile)
EndIf
Sleep(2000)

$myFile = "D:\WorkSpace\Work2016\���Ų����\tools\AutoScrCaptureService\Images\ImageCapture3.jpg"
If FileExists($myFile) Then
    FileDelete($myFile)
EndIf
Sleep(2000)



;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;��ʼ��ͼ

;++++����ͳ������
;ҵ��ͳ��
MouseMove(410,63+100,3);
Sleep(1000);
;ͼ�����
MouseMove(410,176+100,3);
MouseClick($MOUSE_CLICK_LEFT);
MsgBox(0,"��ʾ","�Ͻ�ѡ�����ڶΣ�����20���ӣ�",1);
Sleep(20000);
;ѡ������
MsgBox(0,"��ʾ","ѡ���������������ˣ�",1);
;�����
MouseMove(940,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;ԤԼ��
MouseMove(1123,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;ͳ��
MouseMove(1251,135+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(3000);
;����
MouseMove(344,100+100,3);
Sleep(1000);
MouseMove(1354,624+100,3);
Sleep(1000);
_ScreenCapture_Capture("D:\WorkSpace\Work2016\���Ų����\tools\AutoScrCaptureService\Images\ImageCapture1.jpg",344,100+100,1354,624+100);

;++++�ϼܺ��¼�ͳ������
;ҵ��ͳ��
MouseMove(410,63+100,3);
Sleep(1000);
;ͼ�����¼�
MouseMove(410,156+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;�����
MouseMove(940,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;ԤԼ��
MouseMove(1123,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(1000);
;ͳ��
MouseMove(1251,135+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(3000);
;
MouseMove(344,100+100,3);
Sleep(1000);
MouseMove(1354,624+100,3);
Sleep(1000);
_ScreenCapture_Capture("D:\WorkSpace\Work2016\���Ų����\tools\AutoScrCaptureService\Images\ImageCapture2.jpg",344,100+100,1354,624+100);

;++++��ǰ�ڼ�ͳ������
;ҵ��ͳ��
MouseMove(410,63+100,3);
Sleep(1000);
;��ǰ�ڼ�
MouseMove(410,110+100,3);
MouseClick($MOUSE_CLICK_LEFT);
;ͳ��
MouseMove(1117,122+100,3);
MouseClick($MOUSE_CLICK_LEFT);
Sleep(3000);
;
MouseMove(344,100+100,3);
Sleep(1000);
MouseMove(1354,624+100,3);
Sleep(1000);
_ScreenCapture_Capture("D:\WorkSpace\Work2016\���Ų����\tools\AutoScrCaptureService\Images\ImageCapture3.jpg",344,100+100,1354,624+100);

;++++++++++++++++++++++++++++++++++++++++++++
;��Imges�ļ��У�ѡ�в�����3��ͼƬ�ļ�
RunWait('"' & @ComSpec & '" /c explorer.exe D:\WorkSpace\Work2016\���Ų����\tools\AutoScrCaptureService\Images', '', @SW_HIDE)
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
;΢�����������㿪��Ӧ΢��Ⱥ��Ȼ�󿽱�������
$handle = WinGetHandle("[CLASS:WeChatMainWndForPC]", "");
WinActivate($handle);
Sleep(1000);
$pos = WinGetPos($handle);
MouseMove($pos[0]+100,$pos[1]+30,3);
MouseClick("");
Sleep(1000)
;ControlSend("΢��","","","������ͼ�����");
Send("������ͼ�����");
Sleep(1000);
MouseMove($pos[0]+160,$pos[1]+100,3);
MouseClick("");
Sleep(1000);
;
MouseMove($pos[0]+640,$pos[1]+520,3);
MouseClick("");
;ControlSend("΢��","","","JJJJJJJJJJ{Enter}"); ;OKay
Sleep(1000);
;ControlSend("΢��","","","^v");
Send("^v");
Sleep(1000)
;�㷢��
;ControlSend("΢��","","","{Enter}");
Send("{Enter}");
Sleep(3000)
;������ʾ
;ControlSend("΢��","","","������������ͼ�����в�����ԤԼ����һ�ܽ��ġ����¼ܼ���ǰ�ڼ�ͼ��ͳ�ơ�{Enter}");
Send("������������ͼ�����в�����ԤԼ����һ�ܽ��ġ����¼ܼ���ǰ�ڼ�ͼ��ͳ�ơ�{Enter}");

MsgBox(0,"��ʾ","ȫ����ɣ��ޣ�����",5);

