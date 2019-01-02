function [handles]=imu_calibration(handles)
global maindir
global simdata
settings;
load(strcat(maindir,'\acc_calib.mat'));

path=strcat(maindir,'\����ǰ\');
Files=dir(strcat(path,'*.txt'));
LengthFiles=length(Files);
[m,n]=size(load(strcat(path,Files(1).name)));
imu_data=zeros(m,(n-50)+1);

for i=1:LengthFiles
    set(handles.cali_over,'String',strcat('��',num2str(i),'�����ݿ�ʼ����'));
    pause(1);
    imu_origin_data=load(strcat(path,Files(i).name));
    
    acc=imu_origin_data(:,1:3)'/1365.0;
    gyro=imu_origin_data(:,4:6)'/16.384/180*pi;
    
    for ix=1:length(acc)
        acc(:,ix)= K*(acc(:,ix)-B)*simdata.g;
    end
    imu_data=[-acc(1,50:end)   ; acc(2,50:end); -acc(3,50:end);
              -gyro(1,50:end)  ; gyro(2,50:end);-gyro(3,50:end)];

    cd(maindir(1:(findstr(maindir,'PNS')+2)));
    imu_data(4:6,:)=gyro_bias_calib(imu_data(4:6,:),i);
if ~exist(strcat(maindir,'\������'),'dir')    
    mkdir(strcat(maindir,'\������'));
end
cd(strcat(maindir,'\������'));
save(Files(i).name, '-ASCII','-double', 'imu_data');
set(handles.cali_over,'String',strcat('��',num2str(i),'�����ݲ�������'));
pause(1);
end
set(handles.cali_over,'String','��������!!');

