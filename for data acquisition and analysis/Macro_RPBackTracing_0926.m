%% Original Author: BM %%
% 20220708: Modified by GY (project since 20220615)
% �γ~: �����ѪR20�I
% �γ~: LT RP�^�l
% Last Update: 20230926
% 20230926: direct viewing VA bug fix
close all;clear;clc;
tStartFromBegining = tic;
%% �ϥΪ̿�J
LTID = 31408;
saveLTFile = 0;                                     % �s�� LT File�}��
saveLTRaw = 1;                                      % �s�� LT excel, Fig.�}��
% root folder: will create folder here
rootFolder = "";                                    % ����Ƨ��M�w. "": choose folder; cd: current folder (���Τ޸�); "abc/def...": other path
    customLineFolder = '';                          % ���G��Ƨ��B�~�W��: "MRM5 LT�٭� Custom..."
    customLineFile = '';                            % ���G�ɮ��B�~�W��: "(LT�٭�) II�W Custom..."
GLCritical = 5;                                     % �v���G�ȤƳB�z (�j��ӭȥH�W == 255)

% �������Ѽ� (�C���R������)
% �`�N�����`�Ƥ��o�W�L 4200000 (LT����)
smoothFactor = 0;                                   % �ϥζ}�ҥ���. 0: ����, other: 3,5,7,...21 (���_��)
xOffset = 0;                                        % Receiver LT_X offset (Hor) (���k����)
yOffset = 0;                                        % Receiver LT_Y offset (Ver) (���W����)
zOffset = 0.1;                                      % ���O " ������ "��m (mm)
receiverSizeHor = 165.24;                           % mm (if split on: �`�������j�p)
receiverSizeVer = 293.76;                           % mm (if split on: �`�������j�p)
horGridSize = "";                                   % mm % set to "" if dont needed (if split on: �`�������j�p)
verGridSize = "";                                   % mm % set to "" if dont needed (if split on: �`�������j�p)
horGridNum = 2160;                                  % set to "" if dont needed (if split on: �`�������j�p)
verGridNum = 3840;                                  % set to "" if dont needed (if split on: �`�������j�p)
expectedERR = 0.1;                                  % �w���p�Ȼ~�t % (�V�e�����ȨѰѦ�)
rayFactor = 1;                                      % ���u���B�~�[��

% ���G�v���O�_�վ�j�p
isResultResized = 0;
    resultResizeHorNum = 2160;                      % �վ��v�������V�j�p
    resultResizeVerNum = 3840;                      % �վ��v�������V�j�p

% �O�_�H receiverSize �@���� (�O������j�p)
isReceiverSplit = 1;                                % �нT�O�����ɨS���p���I���ξɦܻ~�t EX: 3840/3 �������~�t
    receiverSplitNumHor = 2;                        % �����V���μƶq
    receiverSplitNumVer = 2;                        % �����V���μƶq

% �����Ѽ� (�M�w������m)
buildLightSouce = 1;                                % �O�_���إ���
moduleTop = 7.62;                                   % �t�γ̳� Z �y�� (mm) 
eyeMode = 0;                                        % -1 0 1 �����k��
pupilSize = 15;                                     % mm
pupilBTDistance = 60;                               % IPD mm
aimAreaHor = 165.24;                                % �����w��d��j�p (����)
aimAreaVer = 293.76;                                % �����w��d��j�p (����)
aimCenterHor = 0;                                   % �����w��d�򤤤� (����)
aimCenterVer = 0;                                   % �����w��d�򤤤� (����)
aimCenterZ = moduleTop;

% ���߳沴 �� VA ��T
WDR = 500;                                          % Based on 0 Offset % (���߲��襤�߭��O)
VVA = 35;                                           % Based on 0 Offset % (���߲��襤�߭��O)
HVA = 0;                                            % Based on 0 Offset % (���߲��襤�߭��O)
systemTiltAngle = 5;                                % �N VA �ഫ�����Ĩ���

% display info for image
dateStringOn = 1;
displayInfo = 1; % VA, STA, PS, eye, IPD
%% ���u�ƹw�� ��L�ѼƽT�{
while (1)
if dateStringOn == 1
    dateString = strcat("_",datestr(now,'mm-dd-yyyy HH-MM'));
elseif dateStringOn == 0
    dateString = "";
end
switch eyeMode;case -1;eyeString = "leftEye";case 0;eyeString = "monoEye";case 1;eyeString = "rightEye";end
imageString = strcat("_WDR",num2str(WDR),"_VVA",num2str(VVA),"_HVA",num2str(HVA),...
    "_STA",num2str(systemTiltAngle),"_PS",num2str(pupilSize),"_",eyeString,"_IPD",num2str(pupilBTDistance));
disp("����:")
disp(imageString)
if displayInfo == 0;imageString = "";end

smoothFactorPool = [0,3,5,7,9,11,13,15,17,19,21];
if ~any(smoothFactor==smoothFactorPool)
    error("���ưѼƳ]�w���~")
end

if ~isequal(customLineFile,"")
    customLineFile = strcat("_",customLineFile);
end

% �Ҽ{ tilt
VVA_Ori = VVA; % ������l���ץ�
HVA_Ori = HVA; % ������l���ץ�
[VVA,HVA] = TiltAngle(WDR,VVA,HVA,systemTiltAngle);

% ���u�� ����� �ˬd
if isReceiverSplit == 1
    receiverSizeHorQuery = receiverSizeHor/receiverSplitNumHor;
    receiverSizeVerQuery = receiverSizeVer/receiverSplitNumVer;
    if isequal(horGridNum,"") && isequal(verGridNum,"")
        actualHGS = horGridSize;
        actualVGS = verGridSize;
        actualHGN = round(receiverSizeHorQuery/horGridSize);
        actualVGN = round(receiverSizeVerQuery/verGridSize); 
    elseif isequal(horGridSize,"") && isequal(verGridSize,"")
        actualHGS = receiverSizeHor/horGridNum;
        actualVGS = receiverSizeVer/verGridNum;
        actualHGN = round(receiverSizeHorQuery/actualHGS);
        actualVGN = round(receiverSizeVerQuery/actualVGS);
    else
        error('�ЫO�� ����� or ����j�p �ܤ@��""');
    end
elseif isReceiverSplit == 0
    receiverSizeHorQuery = receiverSizeHor;
    receiverSizeVerQuery = receiverSizeVer;
    if isequal(horGridNum,"") && isequal(verGridNum,"")
        actualHGS = horGridSize;
        actualVGS = verGridSize;
        actualHGN = round(receiverSizeHorQuery/horGridSize);
        actualVGN = round(receiverSizeVerQuery/verGridSize);
    elseif isequal(horGridSize,"") && isequal(verGridSize,"")
        actualHGS = receiverSizeHorQuery/horGridNum;
        actualVGS = receiverSizeVerQuery/verGridNum;
        actualHGN = horGridNum;
        actualVGN = verGridNum;
    else
        error('�ЫO�� ����� or ����j�p �ܤ@��""');
    end
end

if actualHGN * actualVGN >= 4200000
    error(strcat("���~: ���椤������ (",num2str(horGridNum)," x ",num2str(verGridNum),") �W�L  4200000 �ثe������"))
end

ERR = expectedERR;
rayNum = (actualHGN * actualVGN)/ERR^2;
rayNum = rayNum * rayFactor;
eyePosLT = EyePosition(WDR,VVA,HVA,eyeMode,pupilBTDistance,moduleTop); % �o������m

break
end
%% �Ыظ�Ƨ�
while (1)
if ~isempty(customLineFolder)
    customLineFolder = strcat("_",customLineFolder);
end
targetDirc = strcat("LTRP�^�l ",customLineFolder);  %�۩w�q��Ƨ��W��
rootFolder = string(rootFolder);
if isempty(rootFolder) || isequal(rootFolder,"")
    rootFolder = uigetdir("","��ܥؼи�Ƨ� (�N�b�Ӹ�Ƨ����إ� LT �^�l��Ƨ�)");
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
%% �s��Light Tool (Citrix version)
while (1)
def=System.Reflection.Missing.Value;
% ltcom64path=['C:\Program Files\Optical Research Associates\LightTools 9.1.1\Utilities.NET\LTCOM64.dll'];      %LTCOM64.dll���|
ltcom64path=['C:\Program Files (x86)\Common Files\Optical Research Associates\LightTools\LTCOM64.dll'];      %LTCOM64.dll���|_Critrix!!
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
%% �t��}�l
cprintf("-------------------------------\n")
ReceiverDelete(lt); % �R����e�Ҧ�������
%% �����إ� %%
if buildLightSouce == 1
    checkBLS = 1;
elseif buildLightSouce == 0
    checkBLS = 0;
end
while checkBLS
SourceDelete(lt); % �R����e�Ҧ�Disk���� + ����Cube����
cprintf('key',"�إߥ���......") 
lt.Cmd("\V3D"); % ����@�ɮy�Ф���
% Disk Source �إ�
lt.Cmd('DiskSource ');
lt.Cmd('XYZ'); % �������ߦ�m
lt.Cmd(strcat(num2str(eyePosLT(1)),',',num2str(eyePosLT(2)),',',num2str(eyePosLT(3))));
lt.Cmd('XYZ'); % �b�|��m
lt.Cmd(strcat(num2str(eyePosLT(1)+pupilSize),',',num2str(eyePosLT(2)),',',num2str(eyePosLT(3))));
lt.Cmd('XYZ'); % �k�V�q
lt.Cmd(strcat(num2str(eyePosLT(1)),',',num2str(eyePosLT(2)),',',num2str(eyePosLT(3)-1)));
lt.Cmd('\O"DISK_SOURCE[@Last]"');
lt.Cmd('Name=E');
lt.Cmd('"Power Extent"="Aim region" '); % ���q��w��d��
lt.Cmd('"Aim Entity Type"="Aim Area" '); % �w��ϰ�
lt.Cmd('"Power Units"=Photometric '); % ���q�q
lt.Cmd('\Q');
% ����w��ϰ��}
ObjectList = lt.DbList("DISK_SOURCE[@Last]", "CIRCULAR_AIM_AREA");
ObjectListSize = lt.ListSize(ObjectList); % ������e Source �`��
ObjectKey = lt.ListNext(ObjectList);
ObjectName = lt.DbGet(ObjectKey, "NAME");
lt.ListDelete(ObjectList);
lt.Cmd(strcat('\O"DISK_SOURCE[@Last].CIRCULAR_AIM_AREA[',string(ObjectName),']"'));
lt.Cmd('"Element Shape"=Rectangular '); % �w��ϰ�Ϊ�
lt.Cmd('\Q');
lt.Cmd(strcat('\O"DISK_SOURCE[@Last].RECTANGULAR_AIM_AREA[',string(ObjectName),']"'));
lt.Cmd(strcat('X=',num2str(aimCenterHor))); % �w��ϰ�y�Фj�p
lt.Cmd(strcat('Y=',num2str(aimCenterVer))); % �w��ϰ�y�Фj�p
lt.Cmd(strcat('Z=',num2str(aimCenterZ))); % �w��ϰ�y�Фj�p
lt.Cmd(strcat('Width=',num2str(aimAreaHor))); % �w��ϰ�y�Фj�p
lt.Cmd(strcat('Height=',num2str(aimAreaVer))); % �w��ϰ�y�Фj�p
lt.Cmd('\Q');
cprintf('err',"����\n")
break
end
%% �إ߱�������m�x�} (�|���Ҽ{Offset)
while (1)
if isReceiverSplit == 1
    receiverPositionXYArray = cell(receiverSplitNumVer,receiverSplitNumHor);
    totalNumReceiver = receiverSplitNumVer * receiverSplitNumHor;
    receiverSizeUnitVer = receiverSizeVer/(receiverSplitNumVer*2);
    receiverSizeUnitHor = receiverSizeHor/(receiverSplitNumHor*2);
    positionLeftUp = [-receiverSizeVer/2;-receiverSizeHor/2];
    for ii = 1:receiverSplitNumVer
        for jj = 1:receiverSplitNumHor
            iitemp = (ii-1)*2+1;
            jjtemp = (jj-1)*2+1;
            PositionEach = positionLeftUp + [receiverSizeUnitVer;receiverSizeUnitHor].*[iitemp;jjtemp];               
            receiverPositionXYArray{ii,jj} = [PositionEach(2);-PositionEach(1)]; % ML�y�� > LT�y��
        end
    end
else
    totalNumReceiver = 1;
    receiverSplitNumVer = 1;
    receiverSplitNumHor = 1;
    receiverSizeUnitVer = receiverSizeVer/(receiverSplitNumVer*2);
    receiverSizeUnitHor = receiverSizeHor/(receiverSplitNumHor*2);
    receiverPositionXYArray = cell(receiverSplitNumVer,receiverSplitNumHor);
    receiverPositionXYArray{1,1} = [0;0];
end
break
end
%% excel write check < 218
excelFileNameTest = strcat("(LTRP) R",num2str(totalNumReceiver),customLineFile,".xlsx");
checkExcelFileName(fullfile(fullDirc,excelFileNameTest));
%% �������j��
imageArray = cell(receiverSplitNumVer,receiverSplitNumHor);
for whichReceiver = 1 : totalNumReceiver 
    tStartEachLoop = tic;
    cprintf("-------------------------------\n")
    disp(strcat("��e�B�z:"," WDR",num2str(WDR)," VVA",num2str(VVA_Ori)," HVA",num2str(HVA_Ori)," STA",num2str(systemTiltAngle)," ������ ",num2str(whichReceiver)))
    [row,column] = ind2sub([receiverSplitNumVer,receiverSplitNumHor],whichReceiver);
    receiverPositionXYIncludingOffset = receiverPositionXYArray{row,column} + [xOffset;yOffset];
    rPXYIO = receiverPositionXYIncludingOffset;
    while (1)
    cprintf('key',"�إ߱�����......")
    % Dummy Plane: "D"; Receiver: "R"
%     ReceiverDisable(lt) % �����Τ��e�������� (�o�{�ä��|�O���즳���)
    ReceiverDelete(lt); % �R����e�Ҧ�������
    lt.Cmd("\V3D"); % ����@�ɮy�Ф���
    % dummy plane �إ�
    lt.Cmd('DummyPlane ');
    lt.Cmd('XYZ'); % ���������I
    lt.Cmd(strcat(num2str(rPXYIO(1)),',',num2str(rPXYIO(2)),',',num2str(zOffset-0.001)));
    lt.Cmd('XYZ'); % �������k�u�I (�P���I�۴�k�u)
    lt.Cmd(strcat(num2str(rPXYIO(1)),',',num2str(rPXYIO(2)),',',num2str(zOffset)));
    lt.Cmd('\O"PLANE_DUMMY_SURFACE[@Last]"');
    lt.Cmd(strcat('Name=D',num2str(whichReceiver)));
    lt.Cmd('\Q');
    lt.DbSet("PLANE_DUMMY_SURFACE[@Last]","Width",receiverSizeUnitHor*2);
    lt.DbSet("PLANE_DUMMY_SURFACE[@Last]","Height",receiverSizeUnitVer*2);
    % ������ �إ�
    lt.Cmd(strcat('\O"PLANE_DUMMY_SURFACE[@Last]"'));
    lt.Cmd('"Add Receiver"=');
    lt.Cmd('\Q');
    lt.Cmd('\O"SURFACE_RECEIVER[@Last]"');
    lt.Cmd(strcat('Name=R',num2str(whichReceiver)));
    lt.Cmd('Responsivity=Photometric '); % ��� ���q�q
    lt.Cmd('\Q');
    lt.Cmd('\O"SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward Simulation]"'); % �V�e����
    lt.Cmd('"Has Intensity"=No '); % �����j�פ��R
    lt.Cmd('"Save Ray Data"=No '); % �����x�s���u
    lt.Cmd('\Q');
    lt.Cmd('\O"SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward Simulation].ILLUMINANCE_MESH[Illuminance Mesh]"');
    lt.Cmd(strcat('"X Dimension"=',num2str(actualHGN))); % �]�w����j�pX (Hor)
    lt.Cmd(strcat('"Y Dimension"=',num2str(actualVGN))); % �]�w����j�pY (Ver)
    if smoothFactor == 0
        lt.Cmd('"Do Noise Reduction"=No '); % ��������
    else
        lt.Cmd(strcat('"Kernel Size N"=',num2str(smoothFactor))); % �]�w���Ʊ`��
    end
    lt.Cmd('\Q');
    lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].SIMULATIONS[ForwardAll]" '); % �V�e�����`��
    lt.Cmd('Enabled=Yes '); % �ҥΦV�e����
    lt.Cmd(strcat('MaxProgress=',num2str(rayNum))); % �]�w���u�� (�V�e:���ŦX�~�t����?)
    lt.Cmd('\Q');
    cprintf('err',"����\n")
    break
    end  
    %% �}�l����
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
    errActual = lt.DbGet("SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward_Simulation].ILLUMINANCE_MESH[Illuminance_Mesh]","ErrorAtPeak_Percent");
    cprintf('text',"�p�Ȼ~�t: %.2f %% \n",round(errActual*100)/100)
    %% �s���V�����Raw Datas & LTImage Photos
    while (1)
    if saveLTRaw==1
        cprintf('key',"LT Raw Data �s�ɤ�......")
        meshkey = "SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward_Simulation].ILLUMINANCE_MESH[Illuminance_Mesh]";
        Xdim = lt.DbGet(meshkey, "X_Dimension");
        Ydim = lt.DbGet(meshkey, "Y_Dimension");
        DataArray=zeros(Ydim,Xdim);  %�ʺA�}�C
        [~,DataArray] = lt.GetMeshData(meshkey, DataArray(), "CellValue");
        DataArray=rot90(double(DataArray)); % ���n�B�z
        % �b��
        xAxisFirst = str2double(string(lt.DbGet(meshkey,"XCellCenterAt",def,1)));
        xAxisFinal = str2double(string(lt.DbGet(meshkey,"XCellCenterAt",def,Xdim)));
        xAxisArray = linspace(xAxisFirst,xAxisFinal,Xdim);
        yAxisFirst = str2double(string(lt.DbGet(meshkey,"YCellCenterAt",def,1)));
        yAxisFinal = str2double(string(lt.DbGet(meshkey,"YCellCenterAt",def,Ydim)));
        yAxisArray = linspace(yAxisFirst,yAxisFinal,Ydim)';
        strAtA1 = {'Y\X'};
        % �gExcel: 20220701
        % �ˬd ���ɮ׬O�_�w�g�s�b: if yes --> delete then write
        % �ɮצs�b�U �s��ϥ� 'UseExcel',true �|�XBug 
        excelFileName = strcat("(LTRP) R",num2str(whichReceiver),customLineFile,".xlsx");
        excelFilepath = fullfile(fullDirc,excelFileName);
        if isfile(excelFilepath);delete(excelFilepath);end
        writecell(strAtA1,excelFilepath,'Range','A1','AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Range','B1','AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(yAxisArray,excelFilepath,'Range','A2','AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(DataArray,excelFilepath,'Range','B2','AutoFitWidth',false); % �f�t xlsx2png coding �榡
        cprintf('err',"����\n")
        %% �����
        imageFileName = strcat("(LTRP) (Raw) R",num2str(whichReceiver),customLineFile,".png");
        imageFilepath = fullfile(fullDirc,imageFileName);
        cprintf('key',"�ץX���ɤ�......")
        maxValue = max(max(DataArray));
        minValue = min(min(DataArray));
        LTImage = uint8((DataArray./maxValue)*255);               % 0��̧C��
        if isResultResized == 1
            LTImage = imresize(LTImage,[resultResizeVerNum,resultResizeHorNum]);
        end
        [imageSizeRow,imageSizeColumn] = size(LTImage); % comibine �γ~
        imwrite(LTImage,imageFilepath);
        imageArray{row,column} = LTImage;
        cprintf('err',"����\n")
    end
    break
    end
    cprintf('comment',"Loop %d ����\n",whichReceiver)
    if totalNumReceiver ~= 1
        disp(strcat("��O�ɶ�: ",num2str(toc(tStartEachLoop))," ��"))
    end
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
% �X�����v�� % �HGLCritical�@�G�Ȥ�
while (1)
if isReceiverSplit == 1 && saveLTRaw == 1
    % Raw �v�� (���X)
    imageCombination = ones(actualVGN*receiverSplitNumVer,actualHGN*receiverSplitNumHor);
    for tt = 1:totalNumReceiver
        [row,column] = ind2sub([receiverSplitNumVer,receiverSplitNumHor],tt);
        imageTemp = imageArray{row,column};
        imageCombination(1+actualVGN*(row-1):actualVGN*(row),1+actualHGN*(column-1):actualHGN*(column)) = imageTemp;
    end
    imageCombination = uint8(imageCombination);
    imageCombFileName = strcat("(LTRP)(Raw)(Combination)",dateString,imageString,customLineFile,".png");
    imageCombFilepath = fullfile(fullDirc,imageCombFileName);
    imwrite(imageCombination,imageCombFilepath);
    % �G�ȼv�� (���X)
    imageCombination = ones(actualVGN*receiverSplitNumVer,actualHGN*receiverSplitNumHor);
    for tt = 1:totalNumReceiver
        [row,column] = ind2sub([receiverSplitNumVer,receiverSplitNumHor],tt);
        imageTemp = imageArray{row,column};
        imageTemp(imageTemp>=GLCritical) = 255;
        imageCombination(1+actualVGN*(row-1):actualVGN*(row),1+actualHGN*(column-1):actualHGN*(column)) = imageTemp;
    end
    imageCombination = uint8(imageCombination);
    imageCombFileName = strcat("(LTRP)(GLC",num2str(GLCritical),")(Combination)",dateString,imageString,customLineFile,".png");
    imageCombFilepath = fullfile(fullDirc,imageCombFileName);
    imwrite(imageCombination,imageCombFilepath);
end
break
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
        lt.Cmd('\Q');
    end
    lt.ListDelete(ObjectList);
end
% function ReceiverDisable(lt) % �N���e���������� (�O�d�C��������)
%     lt.Cmd("\V3D");
%     ObjectList = lt.DbList("RECEIVERS[Receiver List]", "SURFACE_RECEIVER");
%     ObjectListSize = lt.ListSize(ObjectList); % ������e Dummy Plane �`��
%     for tt = 1 : ObjectListSize
%         ObjectKey = lt.ListNext(ObjectList);
%         ObjectName = lt.DbGet(ObjectKey, "NAME");
%         ObjectEnabled = lt.DbGet(ObjectKey, "Enabled");
%         if string(ObjectEnabled) == "Yes"
%             lt.Cmd(strcat('\O',"SURFACE_RECEIVER[",string(ObjectName),"]"));
%             lt.Cmd('Enabled=No ');
%             lt.Cmd('\Q');
%         end
%         
%     end
%     lt.ListDelete(ObjectList);
% end
function SourceDelete(lt) % �R����e�Ҧ�Disk Source (���� Cube ����) (�ثe�|�L�k�P�ɧ��Ҧ���������)
    lt.Cmd("\V3D");
    % �R�� DiskSource
    ObjectList = lt.DbList("SOURCES[Source_List]", "DISK_SOURCE");
    ObjectListSize = lt.ListSize(ObjectList); % ������e Source �`��
    for tt = 1 : ObjectListSize
        ObjectKey = lt.ListNext(ObjectList);
        ObjectName = lt.DbGet(ObjectKey, "NAME");
        lt.Cmd(strcat('\ODISK_SOURCE[',string(ObjectName),']'));
        lt.Cmd('Delete');
        lt.Cmd('\Q');
    end
    lt.ListDelete(ObjectList);
    % ���� Cube Source (�������u�i�l��� ����)
    ObjectList = lt.DbList("SOURCES[Source_List]", "CUBE_SURFACE_SOURCE");
    ObjectListSize = lt.ListSize(ObjectList); % ������e Source �`��
    for tt = 1 : ObjectListSize
        ObjectKey = lt.ListNext(ObjectList);
        ObjectName = lt.DbGet(ObjectKey, "NAME");
        lt.Cmd(strcat('\OCUBE_SURFACE_SOURCE[',string(ObjectName),']'));
        lt.Cmd('RayTraceable=No ');
        lt.Cmd('Enabled=No ');
        lt.Cmd('\Q');
    end
    lt.ListDelete(ObjectList);
end
function eyePosLT = EyePosition(WDR,VVA,HVA,eyeMode,pupilBTDistance,moduleTop)
    
    % ���߳沴��m (WDZ: z From systemTop)
    WDZ = WDR*cosd(VVA); 
    viewPointX = WDR*sind(VVA)*cosd(HVA);
    viewPointY = WDR*sind(VVA)*sind(HVA);

    % �Ⲵ���߱���
    eyeVector = [0;eyeMode*pupilBTDistance*0.5]; % ���k�������i�} �S����

    % ���������: no Prism �� HVA �� (HVA=90 && VVA ~=0)
    if HVA~=90
        eyeVector3temp = [eyeVector;0];
        eyeVector3 = rotz(HVA) * eyeVector3temp;
        eyeVector = eyeVector3(1:2);
    end
    eyePos = [viewPointX; viewPointY; moduleTop + WDZ] + [eyeVector;0]; % ML���Шt
    
    % �O Receiver Z = 0 + WD_z; ���O���߬� (x =0 y=0)
    eyePosLT = [eyePos(2);-eyePos(1);eyePos(3)];
    
end
function checkExcelFileName(fullfilename)
    numtest = strlength(fullfilename);
    if numtest > 218
        error("�s�X Excel �����ɦW�r���ƥ����p�� 218")
    end
end
% ��t�ζɱ׮� �����[�ݨ���
function [theta_polar_angle,phi_azimuthal_angle] = TiltAngle(WD,theta_polar_angle,phi_azimuthal_angle,SystemTiltAngle)
    %% 20230927 update
    % ���� VVA ���t
    sign_VVA = abs(theta_polar_angle)/theta_polar_angle;
    
    % �o��ڤH����m
    WD_z=WD*cosd(theta_polar_angle); 
    ViewPoint_x=WD*sind(theta_polar_angle)*cosd(phi_azimuthal_angle);
    ViewPoint_y=WD*sind(theta_polar_angle)*sind(phi_azimuthal_angle);
    pointEye = [ViewPoint_x;ViewPoint_y;WD_z];
    pointEye_roty = roty(-SystemTiltAngle)*pointEye;
    
    % �ϱ� VVA (��������)
    theta_polar_angle = acosd(pointEye_roty(3)/WD);
    
    % �ϱ� HVA
    if pointEye_roty(1)~=0 % �D�����u
        if pointEye_roty(2) == 0 && pointEye_roty(1)< 0 % X �� < 0, HVA 180
                phi_azimuthal_angle = 180;
        elseif pointEye_roty(2) == 0 && pointEye_roty(1)> 0 % X �� > 0, HVA 0
                phi_azimuthal_angle = 0;
        else
            % atand: ��b -90~90 
            phi_azimuthal_angle = atand(pointEye_roty(2)/pointEye_roty(1));
            if pointEye_roty(1) < 0 % atand �|���
               phi_azimuthal_angle = phi_azimuthal_angle + 180;
            end
        end
    else % �b�����u�W�AHVA������ (���t�� VVA ���t�M�w) (EX: no prism DV HVA ��V)
        if sign_VVA < 0
            phi_azimuthal_angle = -phi_azimuthal_angle;
        elseif sign_VVA > 0
%             phi_azimuthal_angle;
        end
    end
    % HVA ����b +- 180 ����
    if phi_azimuthal_angle > 180
        phi_azimuthal_angle = phi_azimuthal_angle - 360;
    elseif phi_azimuthal_angle < -180
        phi_azimuthal_angle = phi_azimuthal_angle + 360;
    end
end