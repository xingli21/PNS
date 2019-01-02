function [] =excel_save(exceldir,imu_data,x_h,cov,zupt,T,foot,group,LengthFiles_R)
%UNTITLED2 Summary of this function goes here
%   Detailed explanation goes here
%�趨Excel�����ļ�����·��
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
sheet_name = strcat(foot,'�ŵ�',num2str(group-(group>LengthFiles_R)*LengthFiles_R),'��');
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
%�ж�Excel�Ƿ��Ѿ��򿪣����Ѵ򿪣����ڴ򿪵�Excel�н��в�����
%����ʹ�Excel
try
    Excel=actxGetRunningServer('Excel.Application');
catch
    Excel = actxserver('Excel.Application');
end;
%����Excel����Ϊ�ɼ�
set(Excel, 'Visible', 1);
%����Excel���������
Workbooks = Excel.Workbooks;
%�������ļ����ڣ��򿪸ñ����ļ��������½�һ���������������棬�ļ���Ϊ����.Excel
if exist(file_report,'file');
    Workbook = invoke(Workbooks,'Open',file_report);
else
    Workbook = invoke(Workbooks, 'Add');
    Workbook.SaveAs(file_report);
end
%���ع�������
Sheets = Excel.ActiveWorkBook.Sheets;
if group>Sheets.Count
    Sheets.Add([],Sheets.Item(Sheets.Count));
end
%���ص�һ�������
% if ~strcmp(Sheets.Item(group).Name,sheet_name)
Sheets.Item(group).Name=sheet_name;
% end
sheet= get(Sheets, 'Item', group);
%�����һ�����
invoke(sheet, 'Activate');
% file_report.Worksheets.Item(group).Name = sheet_name;
%%sheet���õ�����ͷ
sheet_row1={'ʱ��','���ٶȼ�X(m/s^2)','���ٶȼ�Y(m/s^2)','���ٶȼ�Z(m/s^2)','������X��(deg/s)','������Y��(deg/s)','������Z��(deg/s)',...
    'λ��X(m)','λ��Y(m)','λ��Z(m)','�ٶ�X(m/s)','�ٶ�Y(m/s)','�ٶ�Z(m/s)','���(deg)','����(deg)','����(deg)',...
    '������ƫX(deg/s)','������ƫY(deg/s)','������ƫZ(deg/s)','�ӱ���ƫX(m/s^2)','�ӱ���ƫY(m/s^2)','�ӱ���ƫZ(m/s^2)',...
    '���ټ����',....
    '��ˮƽ����','ˮƽλ��������','�ռ�λ��������',...
    '��ʱ��','ZUPT��ʱ��','accX_std','accY_std','accZ_std','gyroX_std','gyroY_std','gyroZ_std',...
    'zupt��ʼ�ٶ�X','zupt��ʼ�ٶ�Y','zupt��ʼ�ٶ�Z','zupt�����ٶ�X','zupt�����ٶ�Y','zupt�����ٶ�Z',...
    'zupt��ʼ�ϳ��ٶ�','zupt�����ϳ��ٶ�'};
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
%     title('�����������')
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
%     hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
%     Excel.ActiveSheet.Range('O1:P1').Select;%��ͼ��ճ������ǰ����A5:B5����
%     Excel.ActiveSheet.Paste;
%     %ɾ��ͼ�ξ��
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
%     title('���߹켣')
%     legend('Trajectory','Start point','End point')
%     xlabel('x [m]')
%     ylabel('y [m]')
%     axis equal
%     grid on
%     box off
%     hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
%     Excel.ActiveSheet.Range('H1:I1').Select;%��ͼ��ճ������ǰ����A5:B5����
%     Excel.ActiveSheet.Paste;
%     %ɾ��ͼ�ξ��
%     delete(zft);
%     zft=figure('visible','off');
%     xl=[1;2;3];
%     bar(xl(1),total_distance(end,1))
%     hold
%     bar(xl(2),Horizontal_error,'r');
%     bar(xl(3),Spherical_error,'g');
%     ylabel('distance(m)')
%     legend('��·��','ˮƽλ��������','λ��������')
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
    d(1,:)=[sheet_row1(1),sheet_row1(10),{'�ϳ��ٶ�(m/s)'},sheet_row1(23)];
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
%     title('zupt�ڼ������������(������)')
%     ylabel('���ٶȼ�(m/s^2)')
%     legend('x-axis','y-axis','z-axis')
%     grid on
%     box off
%     subplot(2,1,2)
%     plot(zupt_gyro_std')
%     ylabel('������(deg/s)')
%     xlabel('zupt����')
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
%     title('zupt��ʼ�ٶ�')
%     ylabel('�ٶ�(m/s)')
%     grid on
%     box off
%     subplot(2,1,2)
%     plot(zupt_vel(4:6,:)')
%     title('zupt�����ٶ�')
%     ylabel('�ٶ�(m/s)')
%     xlabel('zupt����')
%     grid on
%     box off
%     hgexport(zft, '-clipboard');
%     Excel.ActiveSheet.Range('R1:S1').Select;
%     Excel.ActiveSheet.Paste;
%     delete(zft);
%     zft=figure('visible','off');
%     subplot(2,1,1)
%     plot(zupt_velo(1,:)','r')
%     title('zupt��ʼ�ϳ��ٶ�')
%     ylabel('�ٶ�(m/s)')
%     grid on
%     box off
%     subplot(2,1,2)
%     plot(zupt_velo(2,:)','b')
%     title('zupt�����ϳ��ٶ�')
%     ylabel('�ٶ�(m/s)')
%     xlabel('zupt����')
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



