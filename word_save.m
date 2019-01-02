function [] =word_save(worddir,imu_data,x_h,cov,zupt,T,foot,group,LengthFiles_R)
% 设定测试Word文件名和路径 
global simdata;
N=length(cov);
t=0:simdata.Ts:(N-1)*simdata.Ts;
j=1;
t_save=zeros(fix((N-1)/100+1)+1,1);
imu_data_save=zeros(fix((N-1)/100+1)+1,6);
x_h_save=zeros(fix((N-1)/100+1)+1,15);
cov_save=zeros(fix((N-1)/100+1)+1,15);
zupt_save=zeros(fix((N-1)/100+1)+1,1);
for i=1:100:N
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
figure_name = strcat(foot,'脚第',num2str(group-(group>LengthFiles_R)*LengthFiles_R),'组');
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\imu&bias.docx',figure_name,group);
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\position.docx',figure_name,group);
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\height&vel&zupt.docx',figure_name,group);
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\attitude.docx',figure_name,group);
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\covariance.docx',figure_name,group);
saveToWord(exceldir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\zupt_result.docx',sheet_name,group);
end

function []=saveToWord(worddir,time,imu,x,cov,zupt,name,figure_name,group)
file_report = [worddir name];
% 判断Word是否已经打开，若已打开，就在打开的Word中进行操作，否则就打开Word 
try      
% 若Word服务器已经打开，返回其句柄Word     
    Word = actxGetRunningServer('Word.Application'); 
catch      
% 创建一个Microsoft Word服务器，返回句柄Word     
    Word = actxserver('Word.Application');  
end;
% 设置Word属性为可见  
Word.Visible = 1; 
% 若文件存在，打开该文件，否则，新建一个文件，并保存，文件名为filespec_user 
if exist(file_report,'file');       
    Document = Word.Documents.Open(file_report);  % Document = invoke(Word.Documents,'Open',filespec_user); 
else      
    Document = Word.Documents.Add;           % Document = invoke(Word.Documents, 'Add');      
    Document.SaveAs2(file_report); 
end   
Content = Document.Content;   % 返回Content接口句柄 
Selection = Word.Selection;   % 返回Selection接口句柄  
Paragraphformat = Selection.ParagraphFormat;  % 返回ParagraphFormat接口句柄

% 页面设置 
if group==1
    Document.PageSetup.TopMargin = 60;      % 上边距60磅
    Document.PageSetup.BottomMargin = 45;   % 下边距45磅
    Document.PageSetup.LeftMargin = 45;     % 左边距45磅
    Document.PageSetup.RightMargin = 45;    % 右边距45磅
    shape=Document.Shapes;
    shape_count=shape.Count;
    if shape_count~=0;
        for i=1:shape_count;
            shape.Item(1).Delete;
        end;
    end
    Content.Delete;
end

% 设定文档内容的起始位置和标题 
if group~=1
    Selection.InsertBreak;  
end
Selection.Start = Content.end;         % 设置文档内容的起始位置

%如果当前工作表中有图形存在，通过循环将图形全部删除

if strcmp(name,'\imu&bias.docx')
    Selection.Text = figure_name;
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    Selection.Text ='蓝-x,红-y,黄-z';
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off'); %[0.280469 0.553385 0.428906 0.251302]
    subplot(2,1,1)
    plot(time,imu(:,1:3))
    title('惯性器件输出')
    ylabel('Specific force [m/s^2]')
    box off
    grid on
    subplot(2,1,2)
    plot(time,imu(:,4:6)*180/pi)
    xlabel('time [s]')
    ylabel('Angular rate [deg/s]')
    box off
    grid on
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    Selection.Range.PasteSpecial;   % 将图形粘贴到当前文档里
    delete(zft);       % 删除图形句柄
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
 
    zft=figure('Units', 'pixels', 'Position',  [100 100 450 275],'visible','off'); %[0.280469 0.553385 0.428906 0.251302]
    subplot(2,1,1)
    plot(time,x(:,10:12))
    title('Accelerometer bias errors')
    ylabel('Bias [m/s^2]')
    grid on
    box off
    subplot(2,1,2)
    plot(time,x(:,13:15)*180/pi)
    title('Gyroscope bias errors')
    xlabel('time [s]')
    ylabel('Bias [deg/s]')
    box off
    grid on
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    Selection.Range.PasteSpecial;   % 将图形粘贴到当前文档里
    delete(zft);       % 删除图形句柄
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
   
    Document.Save;
    Document.Close;
elseif strcmp(name,'\position.docx')
    Selection.Text = figure_name;
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    zft=figure('Units', 'pixels', 'Position',  [100 100 400 275],'visible','off');
    plot(x(:,2),x(:,1))
    hold
    plot(x(1,2),x(1,1),'rs')
    plot(x(end,2),x(end,1),'bO')
    title('行走轨迹')
    legend('Trajectory','Start point','End point','Orientation','horizontal')
    xlabel('x [m]')
    ylabel('y [m]')
    axis equal
    grid on
    box off
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    Selection.Range.PasteSpecial;   % 将图形粘贴到当前文档里
    delete(zft);       % 删除图形句柄
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
       
    pos=[x(:,2) x(:,1)];
    total_distance=zeros(length(pos),1);
    Horizontal_error=sqrt(sum((x(end,1:2)).^2));
    Spherical_error=sqrt(sum((x(end,1:3)).^2));
    for i=2:length(pos)
        distance= norm(pos(i)-pos(i-1));
        total_distance(i)=total_distance(i-1)+distance;
    end
    zft=figure('Units', 'pixels', 'Position',  [100 100 400 275],'visible','off');
    xl=[1;2;3];
    bar(xl(1),total_distance(end,1))
    hold
    bar(xl(2),Horizontal_error,'r');
    bar(xl(3),Spherical_error,'g');
    ylabel('distance(m)')
    legend('总路程','水平位置最大误差','位置最大误差')
    grid on
    box off
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    Selection.Range.PasteSpecial;   % 将图形粘贴到当前文档里
    delete(zft);       % 删除图形句柄
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
    
    Document.Save;
    Document.Close;
elseif strcmp(name,'\height&vel&zupt.docx')
    Selection.Text = figure_name;
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off');
    subplot(3,1,1)
    plot(time,-x(:,3))
    ylabel('Heigt[m]')
    grid on
    box off
    subplot(3,1,2)
    plot(time,sqrt(sum(x(:,4:6)'.^2))');
    ylabel('Speed [m/s]')
    grid on
    box off
    subplot(3,1,3)
    plot(time,zupt)
    ylabel('Zupt on/off')
    xlabel('time [s]')
    grid on
    box off
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    delete(zft);       % 删除图形句柄
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
    
    Document.Save;
    Document.Close;
elseif strcmp(name,'\attitude.docx')
    Selection.Text = figure_name;
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    Selection.Text ='姿态(蓝-roll,红-pitch,黄-yaw)';
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    zft=figure('Units', 'pixels', 'Position',  [100 100 400 275],'visible','off');
    plot(time,(x(:,7:9))*180/pi)
    title('Attitude')
    xlabel('time [s]')
    ylabel('Angle [deg]')
    grid on
    box off
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    Selection.Range.PasteSpecial;   % 将图形粘贴到当前文档里
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
    
    Document.Save;
    Document.Close;
elseif strcmp(name,'\covariance.docx')
    Selection.Text =figure_name;
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    Selection.Text ='位置/速度(蓝-x,红-y,黄-z),姿态(蓝-roll,红-pitch,黄-yaw)';
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off');
    subplot(3,1,1)
    plot(time,sqrt(cov(:,1:3)))
    title('Position covariance')
    ylabel('[m]')
    grid on
    box off
    subplot(3,1,2)
    plot(time,sqrt(cov(:,4:6)))
    title('Velocity covariance')
    ylabel('[m/s]')
    grid on
    box off
    subplot(3,1,3)
    plot(time,sqrt(cov(:,7:9))*180/pi)
    title('attitude covariance')
    ylabel('[deg]')
    xlabel('time [s]')
    grid on
    box off
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    Selection.Range.PasteSpecial;   % 将图形粘贴到当前文档里
    delete(zft);       % 删除图形句柄
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
    
    Document.Save;
    Document.Close;
elseif strcmp(name,'\zupt_result.docx')
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
    Selection.Text =figure_name;
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    Selection.Text ='蓝-x,红-y,黄-z';
    Selection.Font.Size = 12;   % 设置字号为12
    Selection.Font.Bold = 1;    % 字体不加粗
    Selection.MoveDown;         % 光标下移（取消选中）
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % 左对齐
    Selection.TypeParagraph;          % 回车，另起一段
    
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off');
    subplot(2,1,1)
    plot(zupt_acc_std')
    title('zupt期间惯性器件噪声(均方差)')
    ylabel('加速度计(m/s^2)')
    legend('x-axis','y-axis','z-axis')
    grid on
    box off
    subplot(2,1,2)
    plot(zupt_gyro_std')
    ylabel('陀螺仪(deg/s)')
    xlabel('zupt个数')
    grid on
    box off
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    Selection.Range.PasteSpecial;   % 将图形粘贴到当前文档里
    delete(zft);       % 删除图形句柄
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off');
    subplot(2,1,1)
    plot(zupt_vel(1:3,:)')
    legend('x-axis','y-axis','z-axis')
    title('zupt起始速度')
    ylabel('速度(m/s)')
    grid on
    box off
    subplot(2,1,2)
    plot(zupt_vel(4:6,:)')
    title('zupt结束速度')
    ylabel('速度(m/s)')
    xlabel('zupt个数')
    grid on
    box off
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    Selection.Range.PasteSpecial;   % 将图形粘贴到当前文档里
    delete(zft);       % 删除图形句柄
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off');
    subplot(2,1,1)
    plot(zupt_velo(1,:)','r')
    title('zupt起始合成速度')
    ylabel('速度(m/s)')
    grid on
    box off
    subplot(2,1,2)
    plot(zupt_velo(2,:)','b')
    title('zupt结束合成速度')
    ylabel('速度(m/s)')
    xlabel('zupt个数')
    grid on
    box off
    hgexport(zft, '-clipboard'); %将图形复制到粘贴板
    Selection.Range.PasteSpecial;   % 将图形粘贴到当前文档里
    delete(zft);       % 删除图形句柄
    Selection.MoveRight;           % 光标右移
    Selection.TypeParagraph;          % 回车，另起一段
    
    Document.Save;
    Document.Close;    
end
Word.Visible = 0; 
end