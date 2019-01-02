function [handles]=imu_calibration(handles)
global maindir
global simdata
settings;
load(strcat(maindir,'\acc_calib.mat'));

path=strcat(maindir,'\补偿前\');
Files=dir(strcat(path,'*.txt'));
LengthFiles=length(Files);
[m,n]=size(load(strcat(path,Files(1).name)));
imu_data=zeros(m,(n-50)+1);

for i=1:LengthFiles
    set(handles.cali_over,'String',strcat('第',num2str(i),'组数据开始补偿'));
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
if ~exist(strcat(maindir,'\补偿后'),'dir')    
    mkdir(strcat(maindir,'\补偿后'));
end
cd(strcat(maindir,'\补偿后'));
save(Files(i).name, '-ASCII','-double', 'imu_data');
set(handles.cali_over,'String',strcat('第',num2str(i),'组数据补偿结束'));
pause(1);
end
set(handles.cali_over,'String','补偿结束!!');

