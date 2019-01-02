function varargout = word_report(varargin)
% WORD_REPORT MATLAB code for word_report.fig
%      WORD_REPORT, by itself, creates a new WORD_REPORT or raises the existing
%      singleton*.
%
%      H = WORD_REPORT returns the handle to a new WORD_REPORT or the handle to
%      the existing singleton*.
%
%      WORD_REPORT('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in WORD_REPORT.M with the given input arguments.
%
%      WORD_REPORT('Property','Value',...) creates a new WORD_REPORT or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before word_report_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to word_report_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help word_report

% Last Modified by GUIDE v2.5 31-Dec-2018 00:31:51

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @word_report_OpeningFcn, ...
                   'gui_OutputFcn',  @word_report_OutputFcn, ...
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


% --- Executes just before word_report is made visible.
function word_report_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to word_report (see VARARGIN)

% Choose default command line output for word_report
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes word_report wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = word_report_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on selection change in word_report_list.
function word_report_list_Callback(hObject, eventdata, handles)
% hObject    handle to word_report_list (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns word_report_list contents as cell array
%        contents{get(hObject,'Value')} returns selected item from word_report_list
global maindir
report_file=[maindir '\Word报告'];
try
    % 若Word服务器已经打开，返回其句柄Word
    Word = actxGetRunningServer('Word.Application');
catch
    % 创建一个Microsoft Word服务器，返回句柄Word
    Word = actxserver('Word.Application');
end;
Word.Visible = 1;
if get(handles.word_report_list,'Value')==1
    report_name=[report_file '\imu&bias.docx'];
    if exist(report_name,'file')
        Word.Documents.Open(report_name);
    else
        msgbox('imu&bias.docx报告不存在！','提示框')
    end
end
if get(handles.word_report_list,'Value')==2
    report_name=[report_file '\position.docx'];
    if exist(report_name,'file')
        Word.Documents.Open(report_name);
    else
        msgbox('\position.docx报告不存在！','提示框')
    end
end
if get(handles.word_report_list,'Value')==3
    report_name=[report_file '\height&vel&zupt.docx'];
    if exist(report_name,'file')
        Word.Documents.Open(report_name);
    else
        msgbox('\height&vel&zupt.docx报告不存在！','提示框')
    end
end
if get(handles.word_report_list,'Value')==4
    report_name=[report_file '\attitude.docx'];
    if exist(report_name,'file')
        Word.Documents.Open(report_name);
    else
        msgbox('\attitude.docx报告不存在！','提示框')
    end
end
if get(handles.word_report_list,'Value')==5
    report_name=[report_file '\covariance.docx'];
    if exist(report_name,'file')
        Word.Documents.Open(report_name);
    else
        msgbox('\covariance.docx报告不存在！','提示框')
    end
end


% --- Executes during object creation, after setting all properties.
function word_report_list_CreateFcn(hObject, eventdata, handles)
% hObject    handle to word_report_list (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
