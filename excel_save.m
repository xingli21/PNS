function [] =excel_save(exceldir,imu_data,x_h,cov,zupt,T,foot,group,LengthFiles_R)
%UNTITLED2 Summary of this function goes here
%   Detailed explanation goes here
%设定Excel报告文件名和路径
global simdata;
N=length(cov);
t=0:simdata.Ts:(N-1)*simdata.Ts;
j=1;
% t_save=zeros(fix((N-1)/100+1)+1,1);
% imu_data_save=zeros(fix((N-1)/100+1)+1,6);
% x_h_save=zeros(fix((N-1)/100+1)+1,15);
% cov_save=zeros(fix((N-1)/100+1)+1,15);
% zupt_save=zeros(fix((N-1)/100+1)+1,1);
% for i=1:100:N
%     t_save(j)=t(i);
%     imu_data_save(j,:)=imu_data(:,i)';
%     x_h_save(j,1:15)=x_h(:,i)';
%     cov_save(j,1:15)=cov(:,i)';
%     zupt_save(j,1)=zupt(1,i);
%     j=j+1;
% end
t_save=zeros(N,1);
imu_data_save=zeros(N,6);
x_h_save=zeros(N,15);
cov_save=zeros(N,15);
zupt_save=zeros(N,1);
for i=1:(N-1)
    t_save(j)=t(i);
    imu_data_save(j,:)=imu_data(:,i)';
    x_h_save(j,1:15)=x_h(:,i)';
    cov_save(j,1:15)=cov(:,i)';
    zupt_save(j,1)=zupt(1,i);
    j=j+1;
end
t_save(j)=t(N);
imu_data_save(j,:)=imu_data(:,N)';
x_h_save(j,1:15)=x_h(:,N)';
cov_save(j,1:15)=cov(:,N)';
zupt_save(j,1)=zupt(1,N);
sheet_name = strcat(foot,'脚第',num2str(group-(group>LengthFiles_R)*LengthFiles_R),'组');
saveToExcel(exceldir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\imu&bias.xlsx',sheet_name,group);
saveToExcel(exceldir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\position.xlsx',sheet_name,group);
saveToExcel(exceldir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\height&vel&zupt.xlsx',sheet_name,group);
saveToExcel(exceldir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\attitude.xlsx',sheet_name,group);
saveToExcel(exceldir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\covariance.xlsx',sheet_name,group);
saveToExcel(exceldir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\zupt_result.xlsx',sheet_name,group);
end

function []=saveToExcel(exceldir,time,imu,x,cov,zupt,name,sheet_name,group)
% global simdata;
file_report=[exceldir name];
%判断Excel是否已经打开，若已打开，就在打开的Excel中进行操作，
%否则就打开Excel
try
    Excel=actxGetRunningServer('Excel.Application');
catch
    Excel = actxserver('Excel.Application');
end;
%设置Excel属性为可见
set(Excel, 'Visible', 1);
%返回Excel工作簿句柄
Workbooks = Excel.Workbooks;
%若报告文件存在，打开该报告文件，否则，新建一个工作簿，并保存，文件名为测试.Excel
if exist(file_report,'file');
    Workbook = invoke(Workbooks,'Open',file_report);
else
    Workbook = invoke(Workbooks, 'Add');
    Workbook.SaveAs(file_report);
end
%返回工作表句柄
Sheets = Excel.ActiveWorkBook.Sheets;
if group>Sheets.Count
    Sheets.Add([],Sheets.Item(Sheets.Count));
end
%返回第一个表格句柄
% if ~strcmp(Sheets.Item(group).Name,sheet_name)
Sheets.Item(group).Name=sheet_name;
% end
sheet= get(Sheets, 'Item', group);
%激活第一个表格
invoke(sheet, 'Activate');
% file_report.Worksheets.Item(group).Name = sheet_name;
%%sheet中用到的题头
sheet_row1={'时间','加速度计X(m/s^2)','加速度计Y(m/s^2)','加速度计Z(m/s^2)','陀螺仪X轴(deg/s)','陀螺仪Y轴(deg/s)','陀螺仪Z轴(deg/s)',...
    '位置X(m)','位置Y(m)','位置Z(m)','速度X(m/s)','速度Y(m/s)','速度Z(m/s)','横滚(deg)','俯仰(deg)','航向(deg)',...
    '陀螺零偏X(deg/s)','陀螺零偏Y(deg/s)','陀螺零偏Z(deg/s)','加表零偏X(m/s^2)','加表零偏Y(m/s^2)','加表零偏Z(m/s^2)',...
    '零速检测结果',....
    '总水平距离','水平位置最大误差','空间位置最大误差',...
    '总时间','ZUPT总时间','accX_std','accY_std','accZ_std','gyroX_std','gyroY_std','gyroZ_std',...
    'zupt起始速度X','zupt起始速度Y','zupt起始速度Z','zupt结束速度X','zupt结束速度Y','zupt结束速度Z',...
    'zupt起始合成速度','zupt结束合成速度'};
if strcmp(name,'\imu&bias.xlsx')
%     Shapes=Excel.ActiveSheet.Shapes;
%     if Shapes.Count~=0;
%         for i=1:Shapes.Count;
%             Shapes.Item(1).Delete;
%         end;
%     end;
%     zft=figure('visible','off');%[0.280469 0.553385 0.428906 0.251302]
%     subplot(2,1,1)
%     plot(time,imu(:,1:3))
%     title('惯性器件输出')
%     xlabel('time [s]')
%     ylabel('Specific force [m/s^2]')
%     legend('x-axis','y-axis','z-axis')
%     box off
%     grid on
%     subplot(2,1,2)
%     plot(time,imu(:,4:6)*180/pi)
%     xlabel('time [s]')
%     ylabel('Angular rate  [deg/s]')
%     box off
%     hgexport(zft, '-clipboard'); %将图形复制到粘贴板
%     Excel.ActiveSheet.Range('O1:P1').Select;%将图形粘贴到当前表格的A5:B5栏里
%     Excel.ActiveSheet.Paste;
%     %删除图形句柄
%     delete(zft);
%     zft=figure('visible','off');%[0.280469 0.553385 0.428906 0.251302]
%     subplot(2,1,1)
%     plot(time,x(:,10:12))
%     legend('x-axis','y-axis','z-axis')
%     title('Accelerometer bias errors')
%     xlabel('time [s]')
%     ylabel('Bias [m/s^2]')
%     grid on
%     box off
%     subplot(2,1,2)
%     plot(time,x(:,13:15)*180/pi)
%     title('Gyroscope bias errors')
%     xlabel('time [s]')
%     ylabel('Bias [deg/s]')
%     box off
%     grid on
%     hgexport(zft, '-clipboard');
%     Excel.ActiveSheet.Range('A10:B10').Select;
%     Excel.ActiveSheet.Paste;
%     delete(zft);
%     Workbook.Save;
%     Workbook.Close;
    d=cell(length(time)+1,13);
    d(1,:)=[sheet_row1(1:7),sheet_row1(17:22)];
    for i=1:length(time)
        d(i+1,:)=[num2cell(time(i,:)),num2cell(imu(i,:)),num2cell(x(i,10:15))];
    end
    xlswrite(file_report,d,group);
elseif strcmp(name,'\position.xlsx')
    pos=[x(:,2) x(:,1)];
    total_distance=zeros(length(pos),1);
    Horizontal_error=sqrt(sum((x(end,1:2)).^2));
    Spherical_error=sqrt(sum((x(end,1:3)).^2));
    for i=2:length(pos)
        distance= norm(pos(i)-pos(i-1));
        total_distance(i)=total_distance(i-1)+distance;
    end
%     Shapes=Excel.ActiveSheet.Shapes;
%     if Shapes.Count~=0;
%         for i=1:Shapes.Count;
%             Shapes.Item(1).Delete;
%         end;
%     end;
%     zft=figure('visible','off');
%     plot(x(:,2),x(:,1))
%     hold
%     plot(x(1,2),x(1,1),'rs')
%     plot(x(end,2),x(end,1),'bO')
%     title('行走轨迹')
%     legend('Trajectory','Start point','End point')
%     xlabel('x [m]')
%     ylabel('y [m]')
%     axis equal
%     grid on
%     box off
%     hgexport(zft, '-clipboard'); %将图形复制到粘贴板
%     Excel.ActiveSheet.Range('H1:I1').Select;%将图形粘贴到当前表格的A5:B5栏里
%     Excel.ActiveSheet.Paste;
%     %删除图形句柄
%     delete(zft);
%     zft=figure('visible','off');
%     xl=[1;2;3];
%     bar(xl(1),total_distance(end,1))
%     hold
%     bar(xl(2),Horizontal_error,'r');
%     bar(xl(3),Spherical_error,'g');
%     ylabel('distance(m)')
%     legend('总路程','水平位置最大误差','位置最大误差')
%     grid on
%     box off
%     hgexport(zft, '-clipboard');
%     Excel.ActiveSheet.Range('A10:B10').Select;
%     Excel.ActiveSheet.Paste;
%     delete(zft); 
%     Workbook.Save;
%     Workbook.Close;
    d(1,:)=[sheet_row1(1),sheet_row1(8:9),sheet_row1(24:26)];
    for i=1:length(time)
        d(i+1,:)=[num2cell(time(i,:)),num2cell(x(i,2)),num2cell(x(i,1)),num2cell(total_distance(i,1)),num2cell(Horizontal_error),num2cell(Spherical_error)];
    end
    xlswrite(file_report,d,group);
elseif strcmp(name,'\height&vel&zupt.xlsx')
%     Shapes=Excel.ActiveSheet.Shapes;
%     if Shapes.Count~=0;
%         for i=1:Shapes.Count;
%             Shapes.Item(1).Delete;
%         end;
%     end;
%     zft=figure('visible','off');
%     subplot(3,1,1)
%     plot(time,-x(:,3))
%     xlabel('time [s]')
%     ylabel('Heigt[m]')
%     grid on
%     box off
%     subplot(3,1,2)
%     plot(time,sqrt(sum(x(:,4:6)'.^2))');
%     xlabel('time [s]')
%     ylabel('Speed [m/s]')
%     grid on
%     box off
%     subplot(3,1,3)
%     stem(time,zupt)
%     xlabel('time [s]')
%     ylabel('Zupt on/off')
%     grid on
%     box off
%     hgexport(zft, '-clipboard');
%     Excel.ActiveSheet.Range('F1:G1').Select;
%     Excel.ActiveSheet.Paste;
%     delete(zft);
%     Workbook.Save;
%     Workbook.Close;
    d(1,:)=[sheet_row1(1),sheet_row1(10),{'合成速度(m/s)'},sheet_row1(23)];
    for i=1:length(time)
        d(i+1,:)=[num2cell(time(i,:)),num2cell(x(i,3)),num2cell(sqrt(sum(x(i,4:6).^2))),num2cell(zupt(i,1))];
    end
    xlswrite(file_report,d,group);
elseif strcmp(name,'\attitude.xlsx')
%     Shapes=Excel.ActiveSheet.Shapes;
%     if Shapes.Count~=0;
%         for i=1:Shapes.Count;
%             Shapes.Item(1).Delete;
%         end;
%     end;
%     zft=figure('visible','off');
%     plot(time,(x(:,7:9))*180/pi)
%     title('Attitude')
%     xlabel('time [s]')
%     ylabel('Angle [deg]')
%     legend('Roll','Pitch','Yaw')
%     grid on
%     box off
%     hgexport(zft, '-clipboard');
%     Excel.ActiveSheet.Range('F1:G1').Select;
%     Excel.ActiveSheet.Paste;
%     delete(zft);
%     Workbook.Save;
%     Workbook.Close;
    d(1,:)=[sheet_row1(1),sheet_row1(14:16)];
    for i=1:length(time)
        d(i+1,:)=[num2cell(time(i,:)),num2cell(x(i,7:9)*180/pi)];
    end
    xlswrite(file_report,d,group);  
elseif strcmp(name,'\covariance.xlsx')
%     Shapes=Excel.ActiveSheet.Shapes;
%     if Shapes.Count~=0;
%         for i=1:Shapes.Count;
%             Shapes.Item(1).Delete;
%         end;
%     end;
%     zft=figure('visible','off');
%     subplot(3,1,1)
%     plot(time,sqrt(cov(:,1:3)))
%     title('Position covariance')
%     ylabel('sqrt(cov) [m]')
%     xlabel('time [s]')
%     legend('x-axis', 'y-axis','z-axis')
%     grid on
%     box off
%     subplot(3,1,2)
%     plot(time,sqrt(cov(:,4:6)))
%     title('Velocity covariance')
%     ylabel('sqrt(cov) [m/s]')
%     xlabel('time [s]')
%     grid on
%     box off
%     subplot(3,1,3)
%     plot(time,sqrt(cov(:,7:9))*180/pi)
%     title('attitude covariance')
%     ylabel('sqrt(cov) [deg]')
%     xlabel('time [s]')
%     legend('Roll', 'Pitch','Yaw')
%     grid on
%     box off
%     hgexport(zft, '-clipboard');
%     Excel.ActiveSheet.Range('L1:M1').Select;
%     Excel.ActiveSheet.Paste;
%     delete(zft);
%     Workbook.Save;
%     Workbook.Close;
    d(1,:)=[sheet_row1(1),sheet_row1(8:16)];
    for i=1:length(time)
        d(i+1,:)=[num2cell(time(i,:)),num2cell(sqrt(cov(i,1:6))),num2cell(sqrt(cov(i,7:9))*180/pi)];
    end
    xlswrite(file_report,d,group);
elseif strcmp(name,'\zupt_result.xlsx')
    j=0;
    i=1;
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
%     Shapes=Excel.ActiveSheet.Shapes;
%     if Shapes.Count~=0;
%         for i=1:Shapes.Count;
%             Shapes.Item(1).Delete;
%         end;
%     end;
%     zft=figure('visible','off');
%     subplot(2,1,1)
%     plot(zupt_acc_std')
%     title('zupt期间惯性器件噪声(均方差)')
%     ylabel('加速度计(m/s^2)')
%     legend('x-axis','y-axis','z-axis')
%     grid on
%     box off
%     subplot(2,1,2)
%     plot(zupt_gyro_std')
%     ylabel('陀螺仪(deg/s)')
%     xlabel('zupt个数')
%     grid on
%     box off
%     hgexport(zft, '-clipboard');
%     Excel.ActiveSheet.Range('R1:S1').Select;
%     Excel.ActiveSheet.Paste;
%     delete(zft);
%     zft=figure('visible','off');
%     subplot(2,1,1)
%     plot(zupt_vel(1:3,:)')
%     legend('x-axis','y-axis','z-axis')
%     title('zupt起始速度')
%     ylabel('速度(m/s)')
%     grid on
%     box off
%     subplot(2,1,2)
%     plot(zupt_vel(4:6,:)')
%     title('zupt结束速度')
%     ylabel('速度(m/s)')
%     xlabel('zupt个数')
%     grid on
%     box off
%     hgexport(zft, '-clipboard');
%     Excel.ActiveSheet.Range('R1:S1').Select;
%     Excel.ActiveSheet.Paste;
%     delete(zft);
%     zft=figure('visible','off');
%     subplot(2,1,1)
%     plot(zupt_velo(1,:)','r')
%     title('zupt起始合成速度')
%     ylabel('速度(m/s)')
%     grid on
%     box off
%     subplot(2,1,2)
%     plot(zupt_velo(2,:)','b')
%     title('zupt结束合成速度')
%     ylabel('速度(m/s)')
%     xlabel('zupt个数')
%     grid on
%     box off
%     hgexport(zft, '-clipboard');
%     Excel.ActiveSheet.Range('R1:S1').Select;
%     Excel.ActiveSheet.Paste;
%     delete(zft)
%     Workbook.Save;
%     Workbook.Close;
    d(1,:)=sheet_row1(27:34);
    for i=1:length(time)
        d(i+1,:)=[num2cell(time(end)),num2cell(zupt_t),num2cell(zupt_acc_std(:,i)'),num2cell(zupt_gyro_std(:,i)'),...
                  num2cell(zupt_vel(:,i)'),num2cell(zupt_velo(:,i)')];
    end
    xlswrite(file_report,d,group);
end
end



