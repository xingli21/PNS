function [u]=gyro_bias_calib(u,group)
global maindir
if exist(strcat(maindir,'\gyro_bias.mat'),'file')
    load(strcat(maindir,'\gyro_bias.mat'));
else
    gyro_bias={'组别','零偏最小残差','BiasX','BiasY','BiasZ'};
end
[m,n]=size(gyro_bias);

POINT_FOR_CALIB=8000;
win_size=500;
gyro=u(1:3,:)'*180/pi;
magnitude=zeros(POINT_FOR_CALIB,1);
for i=1:POINT_FOR_CALIB
    first_win=gyro(i:i+win_size,:);
    second_win=gyro(i+win_size:i+2*win_size,:);%%这个地方是否可以改成second_win=gyro(i+1:i+1+win_size,:)，哪个更合适？？邢丽批注
    magnitude(i)=norm(mean(first_win)-mean(second_win));
end
[M,index]=min(magnitude);
bias=mean(gyro(index:index+win_size*2,:))'/180*pi;
gyro_bias(m+1,:)={group M bias(1,1)*180/pi bias(2,1)*180/pi bias(3,1)*180/pi};
for i=1:length(u)
    u(:,i)=u(:,i)-bias;
end
cd(maindir);
save(strcat(maindir,'\gyro_bias.mat'), 'gyro_bias'); %'-ASCII','-double',
end