function [path]=findpath(handles)
global maindir;
global date;
val=get(handles.model_selection,'Value');
currentdir=pwd;%%����Ҫ��PNS�ļ�����
switch val
    case 1
    maindir=strcat(currentdir,'\1#LYX\');
    case 2
    maindir=strcat(currentdir,'\2#HYJ\');
    case 3
    maindir=strcat(currentdir,'\3#XL\');
end
date0=get(handles.test_date,'String');
if length(date0)~=10
    msgbox('����ʱ���ʽ�����⣬����������','��ʾ��');
    path=0;
    set(handles.Word_report_read,'Value',0);
    set(handles.Excel_report_read,'Value',0);
    return;
end
date(1:4)=date0(1:4);
date(5:6)=date0(6:7);
date(7:8)=date0(9:10);
maindir=strcat(maindir,date);
path=1;