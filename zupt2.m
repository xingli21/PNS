function [ zupt_t,zupt_acc_std,zupt_gyro_std,zupt_vel,zupt_velo,t_out] = zupt2( zupt,imu_data,x)
%UNTITLED Summary of this function goes here
%   Detailed explanation goes here
global simdata
N=length(zupt);
t=0:simdata.Ts:(N-1)*simdata.Ts;
j=0;
i=1;
t_out=t(end);
while i<(length(zupt)-1)
    if zupt(i+1)==1
        if zupt(i)==0
            j=j+1;
            index_start(j)=i+1;
        elseif i==1
            j=j+1;
            index_start(j)=i;
        end
        if zupt(i+2)==0
            index_end(j)=i+1;            
        elseif i+2==(length(zupt))
            index_end(j)=i+2;             
        end   
    end
    i=i+1;
end
zupt_t=0;
for i=1:j
    zupt_t=zupt_t+t(index_end(i))-t(index_start(i));
    for m=1:3
        zupt_acc_std(m,i)=std(imu_data(m,index_start(i):index_end(i)));
        zupt_gyro_std(m,i)=std(imu_data(m+3,index_start(i):index_end(i)))*180/pi; 
        zupt_vel(m,i)  = x(m+3,index_start(i));
        zupt_vel(m+3,i)= x(m+3,index_end(i));
    end
    zupt_velo(1,i) = sqrt(sum(x(4:6,index_start(i)).^2));
    zupt_velo(2,i) = sqrt(sum(x(4:6,index_end(i)).^2));
end
figure(100)
subplot(2,1,1)
plot(zupt_acc_std')
title('zupt期间惯性器件噪声(均方差)')
ylabel('加速度计(m/s^2)')
legend('x-axis','y-axis','z-axis')
subplot(2,1,2)
plot(zupt_gyro_std')
ylabel('陀螺仪(deg/s)')
xlabel('zupt个数')

figure(101)
subplot(2,1,1)
plot(zupt_vel(1:3,:)')
legend('x-axis','y-axis','z-axis')
title('zupt起始速度')
ylabel('速度(m/s)')
subplot(2,1,2)
plot(zupt_vel(4:6,:)')
title('zupt结束速度')
ylabel('速度(m/s)')
xlabel('zupt个数')

figure(103)
subplot(2,1,1)
plot(zupt_velo(1,:)','r')
title('zupt起始合成速度')
ylabel('速度(m/s)')
subplot(2,1,2)
plot(zupt_velo(2,:)','b')
title('zupt结束合成速度')
ylabel('速度(m/s)')
xlabel('zupt个数')
end

