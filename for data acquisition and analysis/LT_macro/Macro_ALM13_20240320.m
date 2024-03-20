%% 13 �I���� ���� %%
% Author: GY (project since 20220614) (Ref by BM)
% �T�O Dummy Plane �M Receiver ���R�W (DX RX)
% ���������]�w (�Х��ۦ�]�w����)
% Last Update: 20240320
% content: first git version, 
close all;clear;clc;
tic
%% �ѼƳ]�w
% options
LTID = 13896;           % �P�ɶ}�h��LT�A���w��
saveLTFile = 0;         % �s�� LT File�}��
saveLTRawData = 1;      % �s�� LT excel, Fig.�}��
buildRec = 1;           % �s�� 13 �I Receivers �}�� (0:��������)
chooseFolder = 1;       % ��Ƨ��s��B
customLineFolder = '';    % �ؼи�Ƨ��B�~�R�W
rayFactor = 1;          % ���u���B�~�[��

% parameters setup
% LT�]�w: (�ϦVXY��) �V�kx (HorSize)  �V�Wy (VerSize)
horSize = 165.24;   % For 13 Point Position
verSize = 293.76;   % For 13 Point Position
moduleTop = 0.33;   % mm (�������|�ۦ� = 0.001mm)
% Receiver (type C)
receiverSizeHor = 5; % mm (Also affect 14th Receiver)
receiverSizeVer = 5; % mm (Also affect 14th Receiver)
ALMHalfSize = 2.5; % mm �p�q���j�p�b�| (��ĳ < �������j�p/2)

LgridNum = 360;         % (Also affect 14th Receiver)
VgridNum = 90;          % (Also affect 14th Receiver)
expectedERR = 0.08; % �w���p�Ȼ~�t

MRM = 20;           % �̤j���զ��� (vs Actual Ray Multiplier ARM)

% Debug �� %
NumReceiverArray = 1:13; % 1-13, 14: �ۭq(�w�]������) % 14: ���(�ۭq��m) D38 R38  
    xOffset = 0;         % work when NumReceiverArray = 14
    yOffset = 0;         % work when NumReceiverArray = 14

% excel write �B�z %
dateString = datestr(now,'mm-dd-yyyy HH-MM');
xlsstr = strcat("Result_13�I�V�ᨤ��_MRM ",num2str(MRM),"_",dateString,".xlsx");

%% �Ѽƫe�B�z
while (1)
OOVA = 30; % no use
TotalL = 360;
LgridDegreeStep = TotalL/LgridNum;
L1 = (360 - (LgridNum-1) * LgridDegreeStep) * 0.5;
Lend = L1 + (LgridNum-1) * LgridDegreeStep;
LArray = linspace(L1,Lend,LgridNum);
TotalV = 90;
VgridDegreeStep = TotalV/VgridNum;
V1 = (90 - (VgridNum-1) * VgridDegreeStep) * 0.5;
Vend = V1 + (VgridNum-1) * VgridDegreeStep;
VArray = linspace(V1,Vend,VgridNum);
VArray = fliplr(VArray);
[~,ind] =  min(abs(VArray - OOVA));
VMid = VArray(ind);
Vx = VArray;

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
break
end
%% �Ыظ�Ƨ�
while (1)
if ~isempty(customLineFolder)
    customLineFolder = strcat("_",customLineFolder);
end
targetDirc = strcat("MRM",num2str(MRM)," 13�I����",customLineFolder);  %�۩w�q��Ƨ��W��
if chooseFolder == 0 % �bMatlab��e��Ƨ��إ� 13�I������Ƨ�
    rootDirc = string(cd);
elseif chooseFolder == 1
    rootDirc = uigetdir;
    if rootDirc == 0;return;end
else
    beep
    error("wrong setup for 'chooseFolder'. should be either 0 or 1.")
end
lastwarn('');
fullDirc = fullfile(rootDirc,targetDirc);
mkdir(fullDirc);
msg = lastwarn;
if isequal('Directory already exists.',msg)
    disp("< �۩w�q��Ƨ��w�s�b��e�ؿ� >")
end
excelFilepath = fullfile(fullDirc,xlsstr);
checkExcelFileName(excelFilepath);
break
end
%% �s�� LT (Citrix Version)
while (1)
def=System.Reflection.Missing.Value;
% ltcom64path=['C:\Program Files\Optical Research Associates\LightTools 9.1.1\Utilities.NET\LTCOM64.dll'];      %LTCOM64.dll���|
ltcom64path=['C:\Program Files (x86)\Common Files\Optical Research Associates\LightTools\LTCOM64.dll'];      %LTCOM64.dll���|_Critrix!!
asm=NET.addAssembly(ltcom64path);
lt=LTCOM64.LTAPIx;
lt.LTPID=LTID;
lt.UpdateLTPointer;
lt.Message('�w�s����Matlab����');
break
end
%% 13�I�p�� (+1:��14�I)
while (1)
% LT�]�w: �V�kx (HorSize)  �V�Wy (VerSize)
listPostion = zeros(25,3);
Wr = [1 5 7 8 9 12 13 14 17 18 19 21 25 38];
receiverZ = moduleTop + 0.001;
listPostion(Wr(1),:)=[-9*horSize/20,9*verSize/20,receiverZ];
listPostion(Wr(2),:)=[-9*horSize/20,-9*verSize/20,receiverZ];
listPostion(Wr(3),:)=[-horSize/3,verSize/3,receiverZ];
listPostion(Wr(4),:)=[-horSize/3,0,receiverZ];
listPostion(Wr(5),:)=[-horSize/3,-verSize/3,receiverZ];
listPostion(Wr(6),:)=[0,verSize/3,receiverZ];
listPostion(Wr(7),:)=[0,0,receiverZ];
listPostion(Wr(8),:)=[0,-verSize/3,receiverZ];
listPostion(Wr(9),:)=[horSize/3,verSize/3,receiverZ];
listPostion(Wr(10),:)=[horSize/3,0,receiverZ];
listPostion(Wr(11),:)=[horSize/3,-verSize/3,receiverZ];
listPostion(Wr(12),:)=[9*horSize/20,9*verSize/20,receiverZ];
listPostion(Wr(13),:)=[9*horSize/20,-9*verSize/20,receiverZ];
listPostion(Wr(14),:)=[xOffset,yOffset,receiverZ];
break
end
%% �p�Ȼ~�t�P���u�Ʀ��p
while (1)
ERR = expectedERR;
rayNum = (LgridNum * VgridNum)/ERR^2;
rayNum = rayNum * rayFactor;
break
end
%% Receiver Build
while (1)
if buildRec==1
    cprintf('key',"LT �ҫ��إߤ�......")
    ReceiverDelete(lt);     % �R����e�Ҧ� Dummy Plane
%     sourceName = LightSourceSetup(lt);   % �Ŷ�����]������
    for N = NumReceiverArray
        lt.Cmd("\V3D"); % ����@�ɮy�Ф���
        % dummy plane �إ�
        lt.Cmd('DummyPlane ');
        lt.Cmd('XYZ'); % ���������I
        lt.Cmd(strcat(num2str(listPostion(Wr(N),1)),',',num2str(listPostion(Wr(N),2)),',',num2str(listPostion(Wr(N),3))));
        lt.Cmd('XYZ'); % �������k�u�I (�P���I�۴�k�u)
        lt.Cmd(strcat(num2str(listPostion(Wr(N),1)),',',num2str(listPostion(Wr(N),2)),',',num2str(listPostion(Wr(N),3)+1)));
        lt.Cmd('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"');
        lt.Cmd(strcat('Name=D',num2str(Wr(N))));
        lt.Cmd('\Q');
        lt.DbSet(strcat("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D",num2str(Wr(N)),']'),"Width",receiverSizeHor);
        lt.DbSet(strcat("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D",num2str(Wr(N)),']'),"Height",receiverSizeVer);
        % ������ �إ�
        lt.Cmd(strcat('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"'));
        lt.Cmd('"Add Receiver"=');
        lt.Cmd('\Q');
        lt.Cmd('\O"LENS_MANAGER[1].ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last]"');
        lt.Cmd(strcat('Name=R',num2str(Wr(N))));
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
    end
    cprintf('err',"����\n")
else % (���R��������)
    if ~any(NumReceiverArray == 14) % 13�I�Ҧ�
        ReceiverCheck(lt)
    else
        error("��14�I�Ҧ����䴩�������O�d, �бNbuildRec�]�w��0�p�G�n�i���14�I�Ҧ�")
    end
end
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
if saveLTRawData==1
    cprintf('key',"�g�J��Ƥ�......")
    for N = NumReceiverArray
        meshKey = strcat("SURFACE_RECEIVER[R",num2str(Wr(N)),"].BACKWARD_ANGULAR_LUMINANCE[Angular Luminance].ANGULAR_LUMINANCE_MESH[Angular Luminance Mesh]");
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
        writecell(strAtA1,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','A1','AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','B1','AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(yAxisArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','A2','AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(dataArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','B2','AutoFitWidth',false); % �f�t LT �g�X���% LT Fig (png) �s��
        pngstr = strcat("R",num2str(Wr(N)),".png");
        pngFullFile = fullfile(fullDirc,pngstr);
        lt.Cmd("\V3D");
        lt.Cmd(strcat('LumViewAngularLuminanceChart "R',num2str(Wr(N)),' �V��_����"'));
        lt.Cmd("CopyToClipboard");
        lt.SetOption('ShowFileDialogBox', 0);
        lt.Cmd(strcat("PrintToFile """,pngFullFile,""""));
        lt.SetOption('ShowFileDialogBox', 1);
        lt.Cmd('Dismiss'); % ��������
    end
    cprintf('err',"����\n")
end   
break
end
%% �O���p�Ȼ~�t (20220609 By GY)
while (1)
errorPeakPercentageArray = zeros(5,10);
actualRayHitArray = zeros(5,10);
actualRayMultiplierArray = zeros(5,10);

errorCheck = 0;
for N = NumReceiverArray
    ReceiverMeshKey = strcat("SURFACE_RECEIVER[R",num2str(Wr(N)),"].BACKWARD_ANGULAR_LUMINANCE[Angular Luminance].ANGULAR_LUMINANCE_MESH[Angular Luminance Mesh]");
    [row,column] = ind2sub([5,10],Wr(N));
    errorPeakPercentageArray(row,column) = lt.DbGet(ReceiverMeshKey,"ErrorAtPeak_Percent");
    ReceiverKey = strcat("SURFACE_RECEIVER[R",num2str(Wr(N)),"].BACKWARD_ANGULAR_LUMINANCE[Angular Luminance]");
    actualRayHitArray(row,column) = lt.DbGet(ReceiverKey,"Actual_Ray_Hits");
    actualRayMultiplierArray(row,column) = lt.DbGet(ReceiverKey,"Actual_Ray_Multiplier");
    if errorCheck == 0 && actualRayHitArray(row,column) ~= rayNum
        errorCheck = 1;
        warning("���������S���F�ܳ]�w���u��")
    end
end
if errorCheck == 1
    disp(strcat("max(�̤j���զ���) = ",num2str(max(actualRayMultiplierArray(:)))));
else
    disp("�Ҧ��������F�ܳ]�w���u��.")
end
disp("�p�Ȼ~�t")
disp(errorPeakPercentageArray)
break
end
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
function ReceiverCheck(lt) % �ˬd��e�Ҧ� Dummy Plane �M ������ �R�W�O�_���T
    % �ˬd Dummy Plane
    DummyPlanePool = ["D1","D5","D7","D8","D9","D12","D13","D14","D17","D18","D19","D21","D25"];
    ObjectList = lt.DbList("COMPONENTS[1]", "PLANE_DUMMY_SURFACE");
    ObjectListSize = lt.ListSize(ObjectList); % ������e Dummy Plane �`��
    for tt = 1 : ObjectListSize
        ObjectKey = lt.ListNext(ObjectList);
        ObjectName = string(lt.DbGet(ObjectKey, "NAME"));
        if any(ObjectName == DummyPlanePool)
            DummyPlanePool(any(ObjectName == DummyPlanePool)) = [];
        else
            error("Receiver or Dummy Plane should have certain name rule and amount when buildRec == 0.")
        end
    end
    if ~isempty(DummyPlanePool)
        error("Receiver or Dummy Plane should have certain name rule and amount when buildRec == 0.")
    end
    disp("Dummy Plane �T�{����")
    lt.ListDelete(ObjectList);
    % �ˬd Receiver
    ReceiverPool = ["R1","R5","R7","R8","R9","R12","R13","R14","R17","R18","R19","R21","R25"];
    ObjectList = lt.DbList("RECEIVERS[Receiver_List]", "SURFACE_RECEIVER");
    ObjectListSize = lt.ListSize(ObjectList); % ������e Dummy Plane �`��
    for tt = 1 : ObjectListSize
        ObjectKey = lt.ListNext(ObjectList);
        ObjectName = string(lt.DbGet(ObjectKey, "NAME"));
        if any(ObjectName == ReceiverPool)
            ReceiverPool(any(ObjectName == ReceiverPool)) = [];
        else
            error("Receiver or Dummy Plane should have certain name rule and amount when buildRec == 0.")
        end
    end
    if ~isempty(ReceiverPool)
        error("Receiver or Dummy Plane should have certain name rule and amount when buildRec == 0.")
    end
    disp("Receiver �T�{����")
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
    
    % 2. �Ŷ�����]�w������ (13�IOnly)
    source = strcat("CUBE_SURFACE_SOURCE[",sourceName,"]");
    lt.Cmd("\V3D");
    lt.Cmd(strcat("\O",source,".NATIVE_EMITTER[RightSurface]"));
    lt.Cmd('"Surface Apodizer Type"=Uniform ');
    lt.Cmd('\Q');

    % �̫�M�� object
    lt.ListDelete(ObjectList);
end
function checkExcelFileName(fullfilename)
    numtest = strlength(fullfilename);
    if numtest > 218
        error("�s�X Excel �����ɦW�r���ƥ����p�� 218")
    end
end