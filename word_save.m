function [] =word_save(worddir,imu_data,x_h,cov,zupt,T,foot,group,LengthFiles_R)
% �趨����Word�ļ�����·�� 
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
figure_name = strcat(foot,'�ŵ�',num2str(group-(group>LengthFiles_R)*LengthFiles_R),'��');
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\imu&bias.docx',figure_name,group);
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\position.docx',figure_name,group);
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\height&vel&zupt.docx',figure_name,group);
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\attitude.docx',figure_name,group);
saveToWord(worddir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\covariance.docx',figure_name,group);
saveToWord(exceldir,t_save,imu_data_save,x_h_save,cov_save,zupt_save,'\zupt_result.docx',sheet_name,group);
end

function []=saveToWord(worddir,time,imu,x,cov,zupt,name,figure_name,group)
file_report = [worddir name];
% �ж�Word�Ƿ��Ѿ��򿪣����Ѵ򿪣����ڴ򿪵�Word�н��в���������ʹ�Word 
try      
% ��Word�������Ѿ��򿪣���������Word     
    Word = actxGetRunningServer('Word.Application'); 
catch      
% ����һ��Microsoft Word�����������ؾ��Word     
    Word = actxserver('Word.Application');  
end;
% ����Word����Ϊ�ɼ�  
Word.Visible = 1; 
% ���ļ����ڣ��򿪸��ļ��������½�һ���ļ��������棬�ļ���Ϊfilespec_user 
if exist(file_report,'file');       
    Document = Word.Documents.Open(file_report);  % Document = invoke(Word.Documents,'Open',filespec_user); 
else      
    Document = Word.Documents.Add;           % Document = invoke(Word.Documents, 'Add');      
    Document.SaveAs2(file_report); 
end   
Content = Document.Content;   % ����Content�ӿھ�� 
Selection = Word.Selection;   % ����Selection�ӿھ��  
Paragraphformat = Selection.ParagraphFormat;  % ����ParagraphFormat�ӿھ��

% ҳ������ 
if group==1
    Document.PageSetup.TopMargin = 60;      % �ϱ߾�60��
    Document.PageSetup.BottomMargin = 45;   % �±߾�45��
    Document.PageSetup.LeftMargin = 45;     % ��߾�45��
    Document.PageSetup.RightMargin = 45;    % �ұ߾�45��
    shape=Document.Shapes;
    shape_count=shape.Count;
    if shape_count~=0;
        for i=1:shape_count;
            shape.Item(1).Delete;
        end;
    end
    Content.Delete;
end

% �趨�ĵ����ݵ���ʼλ�úͱ��� 
if group~=1
    Selection.InsertBreak;  
end
Selection.Start = Content.end;         % �����ĵ����ݵ���ʼλ��

%�����ǰ����������ͼ�δ��ڣ�ͨ��ѭ����ͼ��ȫ��ɾ��

if strcmp(name,'\imu&bias.docx')
    Selection.Text = figure_name;
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
    Selection.Text ='��-x,��-y,��-z';
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off'); %[0.280469 0.553385 0.428906 0.251302]
    subplot(2,1,1)
    plot(time,imu(:,1:3))
    title('�����������')
    ylabel('Specific force [m/s^2]')
    box off
    grid on
    subplot(2,1,2)
    plot(time,imu(:,4:6)*180/pi)
    xlabel('time [s]')
    ylabel('Angular rate [deg/s]')
    box off
    grid on
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    Selection.Range.PasteSpecial;   % ��ͼ��ճ������ǰ�ĵ���
    delete(zft);       % ɾ��ͼ�ξ��
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
 
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
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    Selection.Range.PasteSpecial;   % ��ͼ��ճ������ǰ�ĵ���
    delete(zft);       % ɾ��ͼ�ξ��
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
   
    Document.Save;
    Document.Close;
elseif strcmp(name,'\position.docx')
    Selection.Text = figure_name;
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
    zft=figure('Units', 'pixels', 'Position',  [100 100 400 275],'visible','off');
    plot(x(:,2),x(:,1))
    hold
    plot(x(1,2),x(1,1),'rs')
    plot(x(end,2),x(end,1),'bO')
    title('���߹켣')
    legend('Trajectory','Start point','End point','Orientation','horizontal')
    xlabel('x [m]')
    ylabel('y [m]')
    axis equal
    grid on
    box off
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    Selection.Range.PasteSpecial;   % ��ͼ��ճ������ǰ�ĵ���
    delete(zft);       % ɾ��ͼ�ξ��
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
       
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
    legend('��·��','ˮƽλ��������','λ��������')
    grid on
    box off
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    Selection.Range.PasteSpecial;   % ��ͼ��ճ������ǰ�ĵ���
    delete(zft);       % ɾ��ͼ�ξ��
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
    
    Document.Save;
    Document.Close;
elseif strcmp(name,'\height&vel&zupt.docx')
    Selection.Text = figure_name;
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
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
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    delete(zft);       % ɾ��ͼ�ξ��
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
    
    Document.Save;
    Document.Close;
elseif strcmp(name,'\attitude.docx')
    Selection.Text = figure_name;
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
    Selection.Text ='��̬(��-roll,��-pitch,��-yaw)';
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
    zft=figure('Units', 'pixels', 'Position',  [100 100 400 275],'visible','off');
    plot(time,(x(:,7:9))*180/pi)
    title('Attitude')
    xlabel('time [s]')
    ylabel('Angle [deg]')
    grid on
    box off
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    Selection.Range.PasteSpecial;   % ��ͼ��ճ������ǰ�ĵ���
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
    
    Document.Save;
    Document.Close;
elseif strcmp(name,'\covariance.docx')
    Selection.Text =figure_name;
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
    Selection.Text ='λ��/�ٶ�(��-x,��-y,��-z),��̬(��-roll,��-pitch,��-yaw)';
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
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
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    Selection.Range.PasteSpecial;   % ��ͼ��ճ������ǰ�ĵ���
    delete(zft);       % ɾ��ͼ�ξ��
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
    
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
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
    Selection.Text ='��-x,��-y,��-z';
    Selection.Font.Size = 12;   % �����ֺ�Ϊ12
    Selection.Font.Bold = 1;    % ���岻�Ӵ�
    Selection.MoveDown;         % ������ƣ�ȡ��ѡ�У�
    Selection.paragraphformat.Alignment = 'wdAlignParagraphLeft';    % �����
    Selection.TypeParagraph;          % �س�������һ��
    
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off');
    subplot(2,1,1)
    plot(zupt_acc_std')
    title('zupt�ڼ������������(������)')
    ylabel('���ٶȼ�(m/s^2)')
    legend('x-axis','y-axis','z-axis')
    grid on
    box off
    subplot(2,1,2)
    plot(zupt_gyro_std')
    ylabel('������(deg/s)')
    xlabel('zupt����')
    grid on
    box off
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    Selection.Range.PasteSpecial;   % ��ͼ��ճ������ǰ�ĵ���
    delete(zft);       % ɾ��ͼ�ξ��
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off');
    subplot(2,1,1)
    plot(zupt_vel(1:3,:)')
    legend('x-axis','y-axis','z-axis')
    title('zupt��ʼ�ٶ�')
    ylabel('�ٶ�(m/s)')
    grid on
    box off
    subplot(2,1,2)
    plot(zupt_vel(4:6,:)')
    title('zupt�����ٶ�')
    ylabel('�ٶ�(m/s)')
    xlabel('zupt����')
    grid on
    box off
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    Selection.Range.PasteSpecial;   % ��ͼ��ճ������ǰ�ĵ���
    delete(zft);       % ɾ��ͼ�ξ��
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
    zft=figure('Units', 'pixels', 'Position', [100 100 450 275],'visible','off');
    subplot(2,1,1)
    plot(zupt_velo(1,:)','r')
    title('zupt��ʼ�ϳ��ٶ�')
    ylabel('�ٶ�(m/s)')
    grid on
    box off
    subplot(2,1,2)
    plot(zupt_velo(2,:)','b')
    title('zupt�����ϳ��ٶ�')
    ylabel('�ٶ�(m/s)')
    xlabel('zupt����')
    grid on
    box off
    hgexport(zft, '-clipboard'); %��ͼ�θ��Ƶ�ճ����
    Selection.Range.PasteSpecial;   % ��ͼ��ճ������ǰ�ĵ���
    delete(zft);       % ɾ��ͼ�ξ��
    Selection.MoveRight;           % �������
    Selection.TypeParagraph;          % �س�������һ��
    
    Document.Save;
    Document.Close;    
end
Word.Visible = 0; 
end