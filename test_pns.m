function []=test_pns(handles)
global simdata;
global excel_report;
global word_report;
settings;
path=strcat(simdata.path,'\������\');
Files_R=dir(strcat(path,'R-*.txt'));
LengthFiles_R=length(Files_R);
Files_L=dir(strcat(path,'L-*.txt'));
LengthFiles_L=length(Files_L);
Files=[Files_R;Files_L];
LengthFiles=LengthFiles_R+LengthFiles_L;
for i=1:LengthFiles
    if i<=LengthFiles_R
        foot='��';
    else
        foot='��';
    end
    set(handles.test,'String',strcat(foot,'��','��',num2str(i-(i>LengthFiles_R)*LengthFiles_R),'�����ݶ�����...'));
    pause(0.5);
    imu_data=load(strcat(path,Files(i).name)); 
    set(handles.test,'String',strcat(foot,'��','��',num2str(i-(i>LengthFiles_R)*LengthFiles_R),'���������ټ��'));
    pause(0.5);
    [zupt T]=zero_velocity_detector(imu_data);
    set(handles.test,'String',strcat(foot,'��','��',num2str(i-(i>LengthFiles_R)*LengthFiles_R),'�������˲�����'));
    pause(0.5);    
    [x_h cov]=ZUPTaidedINS(imu_data,zupt);
    if excel_report==0
        if ~exist(strcat(simdata.path,'\Excel����'),'dir')
            exceldir=mkdir(strcat(simdata.path,'\Excel����'));
        else
            exceldir=strcat(simdata.path,'\Excel����');
        end
        set(handles.test,'String',strcat(foot,'��','��',num2str(i-(i>LengthFiles_R)*LengthFiles_R),'��Excel���汣����...'));
        pause(0.5);
        excel_save(exceldir,imu_data,x_h,cov,zupt,T,foot,i,LengthFiles_R);
        set(handles.test,'String',strcat(foot,'��','��',num2str(i-(i>LengthFiles_R)*LengthFiles_R),'��Excel�����ѱ��棡'));
        pause(0.5);
    end
    if word_report==0
        if ~exist(strcat(simdata.path,'\Word����'),'dir')
            worddir=mkdir(strcat(simdata.path,'\Word����'));
        else
            worddir=strcat(simdata.path,'\Word����');
        end
        set(handles.test,'String',strcat(foot,'��','��',num2str(i-(i>LengthFiles_R)*LengthFiles_R),'��Word���汣����...'));
        pause(0.5);
        word_save(worddir,imu_data,x_h,cov,zupt,T,foot,i,LengthFiles_R);
        set(handles.test,'String',strcat(foot,'��','��',num2str(i-(i>LengthFiles_R)*LengthFiles_R),'��Word�����ѱ��棡'));
        pause(0.5);
    end

%     word_report(x_h,cov,zupt,T,foot,i);
%     sprintf('Horizontal error = %0.5g , Spherical error = %0.5g',sqrt(sum((x_h(1:2,end)).^2)), sqrt(sum((x_h(1:3,end)).^2)))
end