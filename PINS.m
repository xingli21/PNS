function varargout = PINS(varargin)
% PINS MATLAB code for PINS.fig
%      PINS, by itself, creates a new PINS or raises the existing
%      singleton*.
%
%      H = PINS returns the handle to a new PINS or the handle to
%      the existing singleton*.
%
%      PINS('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PINS.M with the given input arguments.
%
%      PINS('Property','Value',...) creates a new PINS or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before PINS_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to PINS_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help PINS

% Last Modified by GUIDE v2.5 30-Dec-2018 23:24:34

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @PINS_OpeningFcn, ...
                   'gui_OutputFcn',  @PINS_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end

% End initialization code - DO NOT EDIT



% --- Executes just before PINS is made visible.
function PINS_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to PINS (see VARARGIN)

% Choose default command line output for PINS
handles.output = hObject;
% set(handles.start_test,'Enable','Off');
% set(handles.read_report,'Enable','Off');
% global cali_over=1;
set(handles.Word_report_read,'Value',0);
% Update handles structure
guidata(hObject, handles);

% UIWAIT makes PINS wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = PINS_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



% --- Executes on selection change in model_selection.
function model_selection_Callback(hObject, eventdata, handles)
% hObject    handle to model_selection (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% Hints: contents = cellstr(get(hObject,'String')) returns model_selection contents as cell array
%        contents{get(hObject,'Value')} returns selected item from model_selection

% --- Executes during object creation, after setting all properties.
function model_selection_CreateFcn(hObject, eventdata, handles)
% hObject    handle to model_selection (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function test_date_Callback(hObject, eventdata, handles)
% hObject    handle to test_date (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% Hints: get(hObject,'String') returns contents of test_date as text
%        str2double(get(hObject,'String')) returns contents of test_date as a double

% --- Executes during object creation, after setting all properties.
function test_date_CreateFcn(hObject, eventdata, handles)
% hObject    handle to test_date (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on button press in calibration_button.
function calibration_button_Callback(hObject, eventdata, handles)
% hObject    handle to calibration_button (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global maindir;
maindir=[];
global cali_over;
cali_over=0;
if findpath(handles);%获取测试时间所在的文件夹
    if ~exist(maindir,'dir')
        s = sprintf('该测试时间段的数据不存在，请重新确认！');
        msgbox(s,'提示框')
        return
    end
    if ~exist(strcat(maindir,'\补偿后'),'dir')||length(dir(strcat(maindir,'\补偿后')))==2
        cali_over=0;
    else
        button=questdlg('误差已被补偿，是否重新补偿？','提示框','Yes','No','No');
        switch button
            case 'Yes'
                cali_over=0;
            case 'No'
                cali_over=1;
        end
    end
    if cali_over==0
        if ~exist(strcat(maindir,'\补偿前'),'dir')||length(dir(strcat(maindir,'\补偿前')))==2
            msgbox('该数据文件夹为空，请确认','提示框');
            return;
        end
        if exist(strcat(maindir,'\gyro_bias.mat'),'file')
             delete(strcat(maindir,'\gyro_bias.mat'));
        end
        handles=imu_calibration(handles);
        cali_over=1;
    end
    
    if cali_over==1
        set(handles.calibration_button,'Enable','Off');
    end
else 
    return
end
  
% --- Executes on button press in start_test.
function start_test_Callback(hObject, eventdata, handles)
% hObject    handle to start_test (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global maindir;
global excel_report;
global word_report;
if findpath(handles);%获取测试时间所在的文件夹
    if ~exist(maindir,'dir')
        s = sprintf('该测试时间段的数据不存在，请重新确认！');
        msgbox(s,'提示框')
        return
    end
    if ~exist(strcat(maindir,'\补偿后'),'dir')||length(dir(strcat(maindir,'\补偿后')))==2
        msgbox('该测试数据的惯性器件还未进行误差补偿，请先补偿！','提示框')
        return
    end
    if ~exist(strcat(maindir,'\Excel报告'),'dir')||length(dir(strcat(maindir,'\Excel报告')))==2
        excel_report=0;
    else
        button=questdlg('已有Excel报告，是否重新生成？','提示框','Yes','No','No');
        switch button
            case 'Yes'
                excel_report=0;
            case 'No'
                excel_report=1;
        end
    end
    if ~exist(strcat(maindir,'\Word报告'),'dir')||length(dir(strcat(maindir,'\Word报告')))==2
        word_report=0;
    else
        button=questdlg('已有Word报告，是否重新生成？','提示框','Yes','No','No');
        switch button
            case 'Yes'
                word_report=0;
            case 'No'
                word_report=1;
        end
    end
    if excel_report==0||word_report==0
        test_pns(handles);
    end
    set(handles.start_test,'Enable','Off');
else
    return
end

% set(handles.read_report,'Enable','on');    


% --- Executes on button press in radiobutton3.


% --- Executes on button press in Word_report_read.
function Word_report_read_Callback(hObject, eventdata, handles)
% hObject    handle to Word_report_read (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global maindir;
if findpath(handles);%获取测试时间所在的文件夹
    if ~exist(maindir,'dir')
        s = sprintf('该测试时间段的数据不存在，请重新确认！');
        msgbox(s,'提示框')
        set(handles.Word_report_read,'Value',0);
        return
    end
    if ~exist(strcat(maindir,'\Word报告'),'dir')||length(dir(strcat(maindir,'\Word报告')))==2
        msgbox('Word报告还未生成,需要生成！','提示框')
        return
    else
        word_report;
        set(handles.Word_report_read,'Value',0);
    end
else
    set(handles.Word_report_read,'Value',0);
    return
end
% Hint: get(hObject,'Value') returns toggle state of Word_report_read


% --- Executes on button press in Excel_report_read.
function Excel_report_read_Callback(hObject, eventdata, handles)
% hObject    handle to Excel_report_read (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global maindir;
if findpath(handles);%获取测试时间所在的文件夹
    if ~exist(maindir,'dir')
        s = sprintf('该测试时间段的数据不存在，请重新确认！');
        msgbox(s,'提示框')
        set(handles.Excel_report_read,'Value',0);
        return
    end
    if ~exist(strcat(maindir,'\Excel报告'),'dir')||length(dir(strcat(maindir,'\Excel报告')))==2
        msgbox('Excel报告还未生成,需要生成！','提示框')
        return
    else
        excel_report;
    end
else
    return
end
% Hint: get(hObject,'Value') returns toggle state of Excel_report_read
