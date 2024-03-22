%% �V�ᨤ�� ���� %%
% Author: GY (project since 20220614) (Ref by BM)
% Last Update: 20240322
% content: first version
close all;clear;clc
tStartFromBegining = tic;
%% �ѼƳ]�w
% options
LTID = 13896;               % �P�ɶ}�h��LT�A���w��
saveLTFile = 0;             % �s�� LT File �}��
saveLTRaw = 1;              % �s�� LT excel, Fig.�}��

% root folder: will create folder here
rootFolder = "";                                % ����Ƨ��M�w. "": choose folder; cd: current folder (���Τ޸�); "abc/def...": other path
    customLineFolder = '0322';                      % ���G��Ƨ��B�~�W��: "MRM5 LT�٭� Custom..."
    customLineFile = '';                        % ���G�ɮ��B�~�W��: "(LT�٭�) II�W Custom..."

customDirectionGridApodizer = 0;                % �������פ��G (0:Lambertian�A1:�ϥΪ̦ۭq�A-1:�����{�����߰�)
customSurfaceGridApodizer = 1;                  % �����Ŷ����G (0:���áA1:�ϥΪ̦ۭq�A-1:�����{�����߰�)
    % ���k�ɰѼ�
    TXTSwitch = 0;                              % ���k�ɶ}��,�ҥνп�J�U��T�C�Ѽ�
        Zero2One = 1;                           % �O�_�N Matrix ���� 0 �ର 1
        HorSize = 48;                          % (mm)
        VerSize = 27;                           % (mm)
seed = 1;                                       % ���l�ؤl�X�A�i�����N����� (default = 1)
        
% �V�ᨤ�� Receiver �Ѽ� (note: follow "cono" �榡)
smoothFactor = 0;       % �O�_���}����. 0: ����, other: 3,5,7,...21 (���_��)
moduleTop = 0.174;      % mm (�������|�ۦ� = 0.001mm)
xOffset = 0;            % 
yOffset = 0;            % 
receiverSizeHor = 5;    % mm (Also affect 14th Receiver)
receiverSizeVer = 5;    % mm (Also affect 14th Receiver)
ALMHalfSize = 2.5;      % mm �p�q���j�p�b�| (��ĳ < �������j�p/2)
LgridNum = 360;         % (�w�]) ����: �g�� (0-360)
VgridNum = 90;          % (�w�]) ����: �n�� (�W�b�y 0-90)
expectedERR = 0.08;     % �w���p�Ȼ~�t
MRM = 20;               % �̤j���զ��� (vs Actual Ray Multiplier ARM)
rayFactor = 10;         % ���u���B�~�[��
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% �ѼƽT�{
while (1)
% ��L�Ѽ�
dateString = datestr(now,'mm-dd-yyyy HH-MM');

% �p�q�� Check
if ALMHalfSize > min([receiverSizeHor,receiverSizeVer]*0.5)
    beep
    warnAns = questdlg('(Warning) �p�q�����| �j�� �������j�p','Warning','continue','cancel','cancel');
    switch warnAns
        case 'continue'
        otherwise
            disp("System Stopped.")
            return
    end
end
% ���ưѼ� check
smoothFactorPool = [0,3,5,7,9,11,13,15,17,19,21];
if ~any(smoothFactor==smoothFactorPool)
    error("���ưѼƳ]�w���~")
end
% extra_string
if ~isequal(customLineFile,"")
    customLineFile = strcat("_",customLineFile);
end
break
end
%% �Ыظ�Ƨ�
while (1)
if ~isempty(customLineFolder)
    customLineFolder = strcat("_",customLineFolder);
end
targetDirc = strcat("LT �V�ᨤ�� ",customLineFolder);  %�۩w�q��Ƨ��W��
rootFolder = string(rootFolder);
if isempty(rootFolder) || isequal(rootFolder,"")
    rootFolder = uigetdir("","��ܥؼи�Ƨ� (�N�b�Ӹ�Ƨ����إ� LT �V�ᨤ����Ƨ�)");
    if rootFolder == 0;disp("�t�ΰ���");return;end
elseif ~isstring(rootFolder)
    beep
    error("'rootFolder' variable should be either string or char type. (�t�ΰ���)")
end
lastwarn('');
fullDirc = fullfile(rootFolder,targetDirc);
mkdir(fullDirc);
msg = lastwarn;
if isequal('Directory already exists.',msg)
    disp("< �۩w�q��Ƨ��w�s�b��e�ؿ� >")
end
break
end

%% �s�� LT (Citrix Version)
while (1)
def=System.Reflection.Missing.Value;
% ltcom64path='C:\Program Files\Optical Research Associates\LightTools 9.1.1\Utilities.NET\LTCOM64.dll';      %LTCOM64.dll���|
ltcom64path='C:\Program Files (x86)\Common Files\Optical Research Associates\LightTools\LTCOM64.dll';      %LTCOM64.dll���|_Critrix!!
asm=NET.addAssembly(ltcom64path);
lt=LTCOM64.LTAPIx;
lt.LTPID=LTID;
lt.UpdateLTPointer;
lt.Message('�w���\�s��LT (From Matlab)');
% �ˬd status
% [(val),...,status] = lt command
% check whether string(status) == "ltStatusSuccess" or not
% or check status == 0 or not (0 == success)
break
end

%% Ū���ɮ�: ���פ��G, �Ŷ����G�Ҧ��M�w
while (1)
% ���O�����W���^��
sourceName = LightSourceSetup(lt); % ���o��e�����W�� (cube)
% �b����"�k��"�������k��
source = strcat("CUBE_SURFACE_SOURCE[",sourceName,"]");
currentSG_first = string(lt.DbGet(strcat(source,".NATIVE_EMITTER[RightSurface]"),"Surface Apodizer Type"));
currentDG_first = string(lt.DbGet(strcat(source,".NATIVE_EMITTER[RightSurface]"),"Direction Apodizer Type"));
lengthFiles = 1;

%% �������פ��G�]�w
if customDirectionGridApodizer == 1
    cprintf("[info]: �������פ��G������: �ϥΪ̦ۭq ...... ")
    DGType = "Grid";
    % ���פ��G���k��
    pathname_angle=[];
    [filename_angle, pathname_angle] = uigetfile(strcat(pathname_angle,'*.txt'), '�п�ܥ������פ��G���k��');
    if ~ischar(pathname_angle);disp("�t�ΰ���");return;end
    filepath_angle = fullfile(pathname_angle, filename_angle);        % ����/���k�� ��m
elseif customDirectionGridApodizer == 0 
    disp("[info]: �������פ��G������: Lambertian")
    DGType = "Lambertian";
elseif customDirectionGridApodizer == -1 % ����ʷ�e�]�w
    str_temp1 = ["�ϥΪ̦ۭq","Lambertian"];
    if currentDG_first == "Grid"
        str_temp1 = "�ϥΪ̦ۭq";
        DGType = "Grid";
    elseif currentDG_first == "Lambertian"
        str_temp1 = "Lambertian";
        DGType = "Lambertian";
    end
    disp(strcat("[info]: �������פ��G������: ",str_temp1," (�����)"))
else
    error("[error]: �L�k�ѧO�������פ��G���� (customDirectionGridApodizer) (�t�ΰ���)")
end
if customDirectionGridApodizer ~= -1
    lt.Cmd(strcat("\O",source,".NATIVE_EMITTER[RightSurface]")); 
    lt.Cmd(strcat('"Direction Apodizer Type"=',DGType,' ')); 
    if DGType == "Grid" % �פJ���k��
        lt.Cmd("\VConsole "); % �����T����
        lt.Cmd(strcat("\O",source,".NATIVE_EMITTER[RightSurface]"));
        lt.Cmd('"Direction Apodizer Type"=Grid ');
        lt.DbSet(strcat(source,".NATIVE_EMITTER[RightSurface].DIRECTION_GRID_APODIZER[DirectionGridApodizer]"), "LoadFileName", filepath_angle);
        % �ˬd���k�ɦ��ĩ�
        if lt.GetLastMsg(1) == "���~: �פJ���楢�ѡC"
            error("[error]: �L�k�ѧO�������פ��G���k�� (�t�ΰ���)")
        end
        disp("�פJ���k�ɦ��\!")
        lt.Cmd("\V3D "); % ����@�ɮy�Ф���
    end
end
%% �����Ŷ����G�]�w
if customSurfaceGridApodizer == 1
    disp("[info]: �����Ŷ����G������: �ϥΪ̦ۭq")
    SGType = "Grid";
    if TXTSwitch==1 % ��II�� (�۰ʤ��k)
        pathname=[];
        [filename, pathname] = uigetfile(strcat(pathname,'*.png'), '�п�� II �ϧ@�������Ŷ����G���k�� (�i�h��)','MultiSelect','on');
    elseif TXTSwitch==0 % ��II���k��
        pathname=[]; 
        [filename, pathname] = uigetfile(strcat(pathname,'*.txt'), '�п�ܥ����Ŷ����G���k�� (�i�h��)','MultiSelect','on');
    end
    % �p���ɮ׼ƶq
    if ~ischar(pathname) 
        disp("�t�ΰ���");return;end
    if ischar(filename)
        lengthFiles = 1;
    else
        lengthFiles = length(filename);
    end
elseif customSurfaceGridApodizer == 0
    disp("[info]: �����Ŷ����G������: ����")
    SGType = "Uniform";
elseif customSurfaceGridApodizer == -1 % ����ʷ�e�]�w
    str_temp2 = ["�ϥΪ̦ۭq","����"];
    if currentSG_first == "Grid"
        str_temp2 = "�ϥΪ̦ۭq";
        SGType = "Grid";
    elseif currentSG_first == "Uniform"
        str_temp2 = "����";
        SGType = "Uniform";
    end
    disp(strcat("[info]: �����Ŷ����G������: ",str_temp2," (�����)"))
else
    error("[error]: �L�k�ѧO�����Ŷ����G���� (customSurfaceGridApodizer) (�t�ΰ���)")
end
if  customSurfaceGridApodizer ~= -1
    lt.Cmd(strcat("\O",source,".NATIVE_EMITTER[RightSurface]"));
    lt.Cmd(strcat('"Surface Apodizer Type"=',SGType,' '));
end
break
end

%% seed info
seed_firstTime = 1;
disp(strcat("[info]: ���l�ؤl�X: ",num2str(seed)))

%% �p�Ȼ~�t�P���u�Ʀ��p
while (1)
ERR = expectedERR;
rayNum = (LgridNum * VgridNum)/ERR^2;
rayNum = rayNum * rayFactor;
break
end

%% Each File �B�z (�]�t����)
for N = 1:lengthFiles
    tStartEachLoop = tic;
    cprintf("-------------------------------\n")
    %% FullFile
    while (1)
    name = "";
    if customSurfaceGridApodizer == 1
        if lengthFiles == 1
            name = erase(filename,[".png",".bmp",".txt"]);
            filepath = fullfile(pathname, filename);        % ����/���k�� ��m
        else
            name = erase(filename{N},[".png",".bmp",".txt"]);
            filepath = fullfile(pathname, filename{N});     % ����/���k�� ��m
        end
    end
    excelFilepath = fullfile(fullDirc,strcat("(LT ����) ",name,customLineFile,".xlsx"));  % �s���G ��m
    checkExcelFileName(excelFilepath);
    imageFilepath = fullfile(fullDirc,strcat("(LT ����) ",name,customLineFile,".png"));   % �s���G ��m
    break
    end
    
    %% �۰��ন���k�� (if ��J������)
    while (1)
    if SGType == "Grid" % customSurfaceGridApodizer �u�i��O 1 or -1
        if customSurfaceGridApodizer == 1
            if TXTSwitch==1 % ���� ����J
                II = imread(filepath); % valuse: 0-255  
                cprintf('key',"�v���ন���k�ɤ�......")
                func_spa(II,filepath,HorSize,VerSize,Zero2One); % �g�Xtxt��
                filepath = strcat(erase(filepath,[".png",".bmp"]),".txt");
            end
            % �פJ���k��
            lt.DbSet(strcat(source,".NATIVE_EMITTER[RightSurface].SURFACE_GRID_APODIZER[SurfaceGridApodizer]"), "LoadFileName", filepath);
            % �ˬd���k�ɦ��ĩ�
            if lt.GetLastMsg(1) == "���~: �פJ���楢�ѡC"
                disp("[error]: ���~�����k��")
                disp(filepath);
                error("[error]: �L�k�ѧO�����Ŷ����G���k�� (�t�ΰ���)")
            end 
            disp("[info] �Ŷ����k�׶פJ���\!")
        elseif customSurfaceGridApodizer == -1 % �������k��
        end
        % �������k�ɦW
        name = string(lt.DbGet(strcat(source,".NATIVE_EMITTER[RightSurface].SURFACE_GRID_APODIZER[SurfaceGridApodizer]"), "LoadFileName"));
        name = erase(name,".txt");
    end
    break
    end

    %% Receiver Build
    while (1)
    cprintf('key',"LT �ҫ��إߤ�......")
    ReceiverDelete(lt);     % �R����e�Ҧ� Dummy Plane
    sourceName = LightSourceSetup(lt);   % �Ŷ�����]������
    
    lt.Cmd("\V3D"); % ����@�ɮy�Ф���
    % dummy plane �إ�
    lt.Cmd('DummyPlane ');
    lt.Cmd('XYZ'); % ���������I
    lt.Cmd(strcat(num2str(xOffset),',',num2str(yOffset),',',num2str(moduleTop+0.001)));
    lt.Cmd('XYZ'); % �������k�u�I (�P���I�۴�k�u)
    lt.Cmd(strcat(num2str(xOffset),',',num2str(yOffset),',',num2str(moduleTop+0.001+1)));
    lt.Cmd('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"');
    lt.Cmd(strcat('Name=D'));
    lt.Cmd('\Q');
    lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Width",receiverSizeHor);
    lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Height",receiverSizeVer);
    % ������ �إ�
    lt.Cmd(strcat('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"'));
    lt.Cmd('"Add Receiver"=');
    lt.Cmd('\Q');
    lt.Cmd('\O"LENS_MANAGER[1].ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last]"');
    lt.Cmd(strcat('Name=R'));
    lt.Cmd('Responsivity=Photometric '); % ��� ���q�q
    %lt.Cmd('Responsivity=Radiometric '); % ��� ��g�q�q
    lt.Cmd('"Photometry Type"="Photometry Type C" '); % ��V:��������
    lt.Cmd('\Q');
    % ���� �V�e����
    lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward Simulation]"');
    lt.Cmd('Enabled=No ');
    lt.Cmd('"Has Illuminance"=No ');
    lt.Cmd('"Has Intensity"=No ');
    lt.Cmd('\Q');
    % �V����� + ���׽��� �]�w
    lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].BACKWARD[Backward Simulation]"');
    lt.Cmd('"Has Angular Luminance"=Yes ');
    lt.Cmd('\Q');
    lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].BACKWARD_ANGULAR_LUMINANCE[Angular Luminance]"');
    lt.Cmd('ShowRayPreview=No ');
    lt.Cmd('"Save Ray Data"=No ');
    lt.Cmd(strcat('"Max Ray Multiplier"=',num2str(MRM),' '));
    lt.Cmd(strcat('"Ray Hit Goal"=',num2str(rayNum),' '));
    lt.Cmd('\Q');
    lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].BACKWARD_ANGULAR_LUMINANCE[Angular Luminance].ANGULAR_LUMINANCE_MESH[Angular Luminance Mesh]"');
    lt.Cmd(strcat('"X Dimension"=',num2str(LgridNum),' '));
    lt.Cmd(strcat('"Y Dimension"=',num2str(VgridNum),' '));
    lt.Cmd('"Max Y Bound"=90 ');
    lt.Cmd('\Q');
    lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].ANGULAR_LUM_METER[Angular Lum Meter]"');
    lt.Cmd(strcat('HalfSize=',num2str(ALMHalfSize),' '));
    lt.Cmd('\Q');
    
    cprintf('err',"����\n")
    break
    end
    %% Begin Simulation
    while (1)
    cprintf('key',"LT ������......")
    tStart = tic;
    lt.Cmd("\V3D");
    lt.Cmd("BeginAllSimulations");
    tEnd = toc(tStart);
    cprintf('err',"����\n")
    break
    end
    %% �ˬd ���զ���
    errActual = lt.DbGet("SURFACE_RECEIVER[R].BACKWARD_ANGULAR_LUMINANCE[Angular_Luminance].ANGULAR_LUMINANCE_MESH[Angular_Luminance_Mesh]","ErrorAtPeak_Percent");
    cprintf('text',"�p�Ȼ~�t: %.2f %% \n",round(errActual*100)/100)
    ARM = lt.DbGet("SURFACE_RECEIVER[R].BACKWARD_ANGULAR_LUMINANCE[Angular_Luminance]","Actual_Ray_Multiplier");
    if MRM < ARM;warning("�̤j���u���զ��� �p�� ��ڥ��u���զ���");end
    
    %% �s��LT �ɮ�
    while (1)
    if saveLTFile==1
        cprintf('key',"LT �s�ɤ�......")
        backupFileName = strcat("LTFile ",dateString," (Backup)");
        lt.SetOption('ShowFileDialogBox', 0);     % �۰ʦs�� ���s�X�s�ɵ��� (���n)
        lt.Cmd('\VConsole');
        lt.Cmd(strcat("SaveAs """,fullfile(fullDirc,backupFileName),""""));
        lt.SetOption('ShowFileDialogBox', 1);     % �^�_�]�w (���n)
        cprintf('err',"����\n")
    end
    break
    end
    %% �}�ҦV�����-���׽��� (optional)
    % lt.Cmd("\V3D");
    % lt.Cmd(['LumViewAngularLuminanceChart "R13 �V��_����"']);
    %% �s�����
    while (1)
    if saveLTRaw==1
        cprintf('key',"�g�J��Ƥ�......")
        meshKey = strcat("SURFACE_RECEIVER[R].BACKWARD_ANGULAR_LUMINANCE[Angular Luminance].ANGULAR_LUMINANCE_MESH[Angular Luminance Mesh]");
        Xdim = lt.DbGet(meshKey, "X_Dimension");
        Ydim = lt.DbGet(meshKey, "Y_Dimension");
        dataArray=zeros(Ydim,Xdim); 
        [~,dataArray] = lt.GetMeshData(meshKey, dataArray(), "CellValue");
        dataArray=rot90(double(dataArray)); % ���n�B�z
        % �b��
        xAxisFirst = str2double(string(lt.DbGet(meshKey,"XCellCenterAt",def,1)));
        xAxisFinal = str2double(string(lt.DbGet(meshKey,"XCellCenterAt",def,Xdim)));
        xAxisArray = linspace(xAxisFirst,xAxisFinal,Xdim);
        yAxisFirst = str2double(string(lt.DbGet(meshKey,"YCellCenterAt",def,1)));
        yAxisFinal = str2double(string(lt.DbGet(meshKey,"YCellCenterAt",def,Ydim)));
        yAxisArray = linspace(yAxisFirst,yAxisFinal,Ydim)';
        % Raw Data �g�J Excel
        strAtA1 = {'V\L'};
        writecell(strAtA1,excelFilepath,'Sheet',"R",'Range','A1','AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Sheet',"R",'Range','B1','AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(yAxisArray,excelFilepath,'Sheet',"R",'Range','A2','AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(dataArray,excelFilepath,'Sheet',"R",'Range','B2','AutoFitWidth',false); % �f�t LT �g�X���
        % LT Fig (png) �s��
        lt.Cmd("\V3D");
        lt.Cmd(strcat('LumViewAngularLuminanceChart "R �V��_����"'));
        lt.Cmd("CopyToClipboard");
        lt.SetOption('ShowFileDialogBox', 0);
        lt.Cmd(strcat("PrintToFile """,imageFilepath,""""));
        lt.SetOption('ShowFileDialogBox', 1);
        lt.Cmd('Dismiss'); % ��������
        cprintf('err',"����\n")
    end   
    break
    end
    cprintf('comment',"Loop����\n")
    if lengthFiles ~= 1
        disp(strcat("��O�ɶ�: ",num2str(toc(tStartEachLoop))," ��"))
    end
    % ���� flush memory
    lt.Cmd("FlushUndoMemory");
    lt.Cmd("FlushDeletedEntityMemory");
    lt.Cmd("FlushSavedRayDataMemory");
    lt.Cmd("FlushAllMemory");
    pause(2)
end
cprintf('text',"�{�ǧ���\n")
disp(strcat("�`��O�ɶ�: ",num2str(toc(tStartFromBegining))," ��"))

%% Function
function ReceiverDelete(lt) % �R����e�Ҧ�DummyPlane (�Y�]�t������)
    lt.Cmd("\V3D");
    ObjectList = lt.DbList("COMPONENTS[1]", "PLANE_DUMMY_SURFACE");
    ObjectListSize = lt.ListSize(ObjectList); % ������e Dummy Plane �`��
    for tt = 1 : ObjectListSize
        ObjectKey = lt.ListNext(ObjectList);
        ObjectName = lt.DbGet(ObjectKey, "NAME");
        lt.Cmd(strcat('\O',"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[",string(ObjectName),']'));
        lt.Cmd('Delete');
    end
    lt.ListDelete(ObjectList);
end
function sourceName = LightSourceSetup(lt) % �����]�w
    % �ثe�\��:
    % 1. �ˬd���ĥ��� (cube)
    % 2. �Ŷ�����אּ����
    %%%%%%
    % 1. �ˬd����
    lt.Cmd("\V3D");
    ObjectList = lt.DbList("SOURCES[Source_List]", "CUBE_SURFACE_SOURCE");
    ObjectListSize = lt.ListSize(ObjectList); % ������e LightSource (Cube) �`��
    if ObjectListSize == 0
        error("���~: ��e LT Model �S������ (cube)")
    end
    sucCheck = 0;
    countAvailable = 0;
    for tt = 1 : ObjectListSize
        ObjectKey = lt.ListNext(ObjectList);
        if lt.DbGet(ObjectKey, "RayTraceable") == "Yes"...
                && lt.DbGet(ObjectKey, "Enabled") == "Yes"
            sourceName = string(lt.DbGet(ObjectKey, "NAME"));
            sucCheck = 1;
        end
        if lt.DbGet(ObjectKey, "RayTraceable") == "Yes"...
                || lt.DbGet(ObjectKey, "Enabled") == "Yes"
            countAvailable = countAvailable + 1;
        end
    end
    if sucCheck == 0
        error("���~: �䤣��i�Ϊ�����")
    end
    if countAvailable ~= 1
        error("���~: ���D����, ��L�����г]�w 1. �������u�i�l��� 2. ����")
    end

    % �̫�M�� object
    lt.ListDelete(ObjectList);
end
function checkExcelFileName(fullfilename)
    numtest = strlength(fullfilename);
    if numtest > 218
        error("�s�X Excel �����ɦW�r���ƥ����p�� 218")
    end
end