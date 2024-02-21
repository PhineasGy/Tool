%% Original Author: BM %%
% 20220620: Modified by GY (project since 20220620)
% �γ~: 13 �I��]�P���R
% �T�O Dummy Plane �M Receiver ���R�W (DX RX)
% �T�O ���B�u���@�� �ҥ�+���u�i�l��ʶ}�� �� CUBE_SURFACE_SOURCE
% Last Update: 20230926
% content: direct viewing VA bug fix
close all;clear;clc;
tStartFromBegining = tic;
%% �ϥΪ̿�J
LTID = 19460;
saveLTFile = 0;                                     % �s�� LT File�}��
saveLTRaw = 1;                                      % �s�� LT excel, Fig.�}��
chooseFolder = 1;                                   % ���root��Ƨ� (�|�b�Ӹ�Ƨ��s��LT�٭��Ƨ�)
    customLineFolder = '';                          % ��Ƨ��W��: "MRM5 LT�٭� Custom..."
    customLineFile = '';                            % RawData/Png�W��: "(LT�٭�) II�W Custom..."
customDirectionGridApodizer = 1;                    % �������פ��G (0:Lambertian�A1:�ϥΪ̦ۭq�A-1:�����{�����߰�)
seedPool = 1:5;                                     % random example: randi([1,100],[1,3]) ���T�� 1-100 �� seed
NumReceiverArray = 1:13;                            % 1-13

% 13 �I��m�]�w
% LT�]�w: (�ϦVXY��) �V�kx (HorSize)  �V�Wy (VerSize)
moduleHorSize = 165.24;     % mm
moduleVerSize = 293.76;     % mm
moduleTop = 11.99;          % mm

% Receiver
buildRec = 1;                                       % �s��13�IReceivers�}�� (0:������)
    smoothFactor = 0;                               % �O�_���}����. 0: ����, other: 3,5,7,...21 (���_��)
    receiverSizeHor = 5; % mm
    receiverSizeVer = 5; % mm
    humanFactor = 1;
    verGridSize = 0.0765; % mm
    horGridSize = 0.0765; % mm
    horGridNum = "";
    verGridNum = "";
expectedERR = 0.02;
MRM = 20;                                           % �̤j���զ��� (vs Actual Ray Multiplier ARM)
rayFactor = 1;                                      % ���u���B�~�[��

% �����Ѽ� (�V��Ž�)
eyeMode = 0;                                        % -1 0 1 �����k��
pupilSize = 2.5;                                    % (mm)
pupilBTDistance = 60;                               % IPD (mm)
% VA Array % ��� ���߲� �� ���߭��O % �i�]�w PL_Array
viewingSetWDR = [400];
viewingSetVVA = [30];
viewingSetHVA = [0];                                
systemTiltAngle = 0;    

% Erosion Setup
erosionCS= 1;
structureElementRadius = 3;                         % 1 mm �N��X�� Pixel (TBD)
structureElementHeight = 0;                         % offset
%% PreProcessing
while (1)
saveLTRawData = 1;     % �����}��
% �ˬd WDR VVA HVA �ռ�
if ~(length(viewingSetWDR)==length(viewingSetVVA)&&length(viewingSetWDR)==length(viewingSetHVA))
    error("VA�ռƥ����ۦP");
end
dateString = datestr(now,'mm-dd-yyyy HH-MM');%% structure element
% erosion unit
se = offsetstrel('ball',structureElementRadius, structureElementHeight);
break
end
%% �]�Ʋյ��G Loop
onlyCheckFirstTime = 0;
seedNum = length(seedPool);
setNum = length(viewingSetWDR);
CSMeanVAArray = cell(1,setNum);
for whichVANum = 1:setNum
tStartVASet = tic;
intensityArrayEachSeed = cell(1,seedNum); % ��l��
CSArrayEachSeed = cell(1,seedNum);
close all
WDR = viewingSetWDR(whichVANum);
VVA = viewingSetVVA(whichVANum);
HVA = viewingSetHVA(whichVANum);
% �Ҽ{ tilt
VVA_Ori = VVA; % ������l���ץ�
HVA_Ori = HVA; % ������l���ץ�
[VVA,HVA] = TiltAngle(WDR,VVA,HVA,systemTiltAngle);
disp(strcat("�B�z��: ","WD",num2str(WDR)," VVA",num2str(VVA_Ori)," HVA",num2str(HVA_Ori)," STA",num2str(systemTiltAngle)))
%% seed Pool
for whichSeedNum = 1:seedNum
disp("------------")
tStartEachSeed = tic;
disp(strcat("Seed: ",num2str(seedPool(whichSeedNum))))
whichSeed = seedPool(whichSeedNum);  
onlyCheckFirstTime = onlyCheckFirstTime + 1;
%% �Ыظ�Ƨ� (�u�b�Ĥ@���إ�)
while (1)
if onlyCheckFirstTime == 1
    if ~isempty(customLineFolder)
        customLineFolder = strcat(" ",customLineFolder);
    end
end
targetDirc = strcat("MRM",num2str(MRM)," 13�I�Ž� Seed ",num2str(whichSeed),customLineFolder);  %�۩w�q��Ƨ��W��
rootFolder = string(rootFolder);
if isempty(rootFolder) || isequal(rootFolder,"")
    if onlyCheckFirstTime == 1 % �u��@����Ƨ���l
        rootFolder = uigetdir("","��ܥؼи�Ƨ� (�N�b�Ӹ�Ƨ����إ�LT�٭��Ƨ�)");
        if rootFolder == 0;disp("�t�ΰ���");return;end
    end
elseif ~isstring(rootFolder)
    beep
    error("'rootFolder' variable should be either string or char type. (�t�ΰ���)")
end

lastwarn(''); % ���m warning
fullDirc = fullfile(rootFolder,targetDirc);
mkdir(fullDirc);
msg = lastwarn;
if isequal('Directory already exists.',msg)
    disp("< �۩w�q��Ƨ��w�s�b��e�ؿ� >")
end
break
end
%% �s�� LT (Citrix Version) % �u�b�Ĥ@���إ�
while (onlyCheckFirstTime==1) 
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
%% seed
lt.Cmd("\V3D"); % ����@�ɮy�Ф���
lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].SIMULATIONS[ForwardAll]" ');
lt.Cmd(strcat("StartingSeed=",num2str(whichSeed)));
lt.Cmd('\Q');
%%
%% Ū���ɮ�: ���פ��G�Ҧ��M�w (�u�B�z�@��)
while onlyCheckFirstTime == 1
% ���O�����W���^��
sourceName = LightSourceSetup(lt); % ���o��e�����W�� (cube)
% �b����"�k��"�������k��
source = strcat("CUBE_SURFACE_SOURCE[",sourceName,"]");
currentDG_first = string(lt.DbGet(strcat(source,".NATIVE_EMITTER[RightSurface]"),"Direction Apodizer Type"));
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
break
end
%% 13�I�p��
while (1)
% LT�]�w: �V�kx (HorSize)  �V�Wy (VerSize)
receiverPosition = zeros(25,3);
Wr = [1 5 7 8 9 12 13 14 17 18 19 21 25];
receiverZ = moduleTop + 0.001;
viewPointX = WDR*sind(VVA)*cosd(HVA); % ������VVA
viewPointY = WDR*sind(VVA)*sind(HVA); % ������VVA
pupilCenterXY = [0;eyeMode*pupilBTDistance*0.5]; %�H���y���߬��s�I
% �Ⲵ����
% ���������: no Prism ��HVA�� (HVA=90 && VVA ~=0)
if HVA ~= 90
    pupilCenterXYZTemp=[pupilCenterXY;0];
    pupilCenterXYZRotTemp = rotz(HVA) * pupilCenterXYZTemp;
    pupilCenterXY = pupilCenterXYZRotTemp(1:2); % �����V�q
else
    pupilCenterXY = [0;0];
end
eyePosition = [WDR*sind(VVA)*sind(HVA),-WDR*sind(VVA)*cosd(HVA),receiverZ+WDR*cosd(VVA)]...
    + [pupilCenterXY(2),-pupilCenterXY(1),0];
receiverPosition(Wr(1),:)=[-9*moduleHorSize/20,9*moduleVerSize/20,receiverZ];
receiverPosition(Wr(2),:)=[-9*moduleHorSize/20,-9*moduleVerSize/20,receiverZ];
receiverPosition(Wr(3),:)=[-moduleHorSize/3,moduleVerSize/3,receiverZ];
receiverPosition(Wr(4),:)=[-moduleHorSize/3,0,receiverZ];
receiverPosition(Wr(5),:)=[-moduleHorSize/3,-moduleVerSize/3,receiverZ];
receiverPosition(Wr(6),:)=[0,moduleVerSize/3,receiverZ];
receiverPosition(Wr(7),:)=[0,0,receiverZ];
receiverPosition(Wr(8),:)=[0,-moduleVerSize/3,receiverZ];
receiverPosition(Wr(9),:)=[moduleHorSize/3,moduleVerSize/3,receiverZ];
receiverPosition(Wr(10),:)=[moduleHorSize/3,0,receiverZ];
receiverPosition(Wr(11),:)=[moduleHorSize/3,-moduleVerSize/3,receiverZ];
receiverPosition(Wr(12),:)=[9*moduleHorSize/20,9*moduleVerSize/20,receiverZ];
receiverPosition(Wr(13),:)=[9*moduleHorSize/20,-9*moduleVerSize/20,receiverZ];
incidentVector = eyePosition - receiverPosition;
break
end
%% Get LTWDR LTVVA LTHVA for SLM Input (listAngleForSLM) (ToBeChecked)
while (1)
VALTList = nan(25,3); % wdr,theta,phi
P2A = @Position2Angle;
for jj = Wr
    [wdrEach,thetaEach,phiEach] = P2A(-incidentVector(jj,2),incidentVector(jj,1),incidentVector(jj,3));
    VALTList(jj,1) = wdrEach;
    VALTList(jj,2) = thetaEach;
    VALTList(jj,3) = phiEach;
end
VALTList(:,3)= mod(360-VALTList(:,3),360);
break
end
%% �H�]�]�w ���u�Ƴ]�w
while (1)
if humanFactor == 1
    angRes = tand(1/120); % �H��������v
    HFGridSize = WDR * angRes * 2;
    actualHGS = HFGridSize;
    actualVGS = HFGridSize;
    actualHGN = round(receiverSizeHor/actualHGS);
    actualVGN = round(receiverSizeVer/actualVGS);
elseif humanFactor == 0
    if horGridNum == "" && verGridNum == ""
        actualHGS = horGridSize;
        actualVGS = verGridSize;
        actualHGN = round(receiverSizeHor/horGridSize);
        actualVGN = round(receiverSizeVer/verGridSize);
    elseif horGridSize == "" && verGridSize == ""
        actualHGS = receiverSizeHor/horGridNum;
        actualVGS = receiverSizeVer/verGridNum;
        actualHGN = horGridNum;
        actualVGN = verGridNum;
    else
        if round(receiverSizeHor/horGridSize) ~=  horGridNum ||...
                round(receiverSizeVer/verGridSize) ~=  verGridNum
            warning("Mesh �Ƴ]�w�i�঳�~: �ثe�H horGridSize / verGridSize ���D���� horGridNum / verGridNum")
            actualHGS = horGridSize;
            actualVGS = verGridSize;
            actualHGN = round(receiverSizeHor/horGridSize);
            actualVGN = round(receiverSizeVer/verGridSize);
        else
            actualHGN = horGridNum;
            actualVGN = verGridNum;
            actualHGS = horGridSize;
            actualVGS = verGridSize;
        end
    end
end
ERR = expectedERR;
rayNum = (actualHGN * actualVGN)/ERR^2;
rayNum = rayNum * rayFactor;
break
end
%% Receiver Build
while (1)
cprintf('key',"LT �ҫ��إߤ�......")
ReceiverDelete(lt); % �R����e�Ҧ� Dummy Plane
sourceName = LightSourceSetup(lt); % �Ŷ�����]������
if buildRec==1
    for N = NumReceiverArray
        WDRLT = VALTList(Wr(N),1);
        VVALT = VALTList(Wr(N),2);
        HVALT = VALTList(Wr(N),3);
        lt.Cmd("\V3D"); % ����@�ɮy�Ф���
        % dummy plane �إ�
        lt.Cmd('DummyPlane ');
        lt.Cmd('XYZ'); % ���������I
        lt.Cmd(strcat(num2str(receiverPosition(Wr(N),1)),',',num2str(receiverPosition(Wr(N),2)),',',num2str(receiverPosition(Wr(N),3))));
        lt.Cmd('XYZ'); % �������k�u�I (�P���I�۴�k�u)
        lt.Cmd(strcat(num2str(receiverPosition(Wr(N),1)),',',num2str(receiverPosition(Wr(N),2)),',',num2str(receiverPosition(Wr(N),3)+1)));
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
        lt.Cmd('Responsivity=Photometric ');
        lt.Cmd('\Q');
        % ���� �V�e����
        lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward Simulation]"');
        lt.Cmd('Enabled=No ');
        lt.Cmd('"Has Illuminance"=No ');
        lt.Cmd('"Has Intensity"=No ');
        lt.Cmd('\Q');
        % �V����� + �Ŷ����� �]�w
        lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].BACKWARD[Backward Simulation]"');
        lt.Cmd('"Has Spatial Luminance"=Yes ');
        lt.Cmd('\Q');
        lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].BACKWARD_SPATIAL_LUMINANCE[Spatial Luminance]"');
        lt.Cmd('ShowRayPreview=No ');
        lt.Cmd('"Save Ray Data"=No ');
        lt.Cmd(strcat('"Max Ray Multiplier"=',num2str(MRM),' '));
        lt.Cmd(strcat('"Ray Hit Goal"=',num2str(rayNum),' '));
        lt.Cmd('\Q');
        lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].BACKWARD_SPATIAL_LUMINANCE[Spatial Luminance].SPATIAL_LUMINANCE_MESH[Spatial Luminance Mesh]"');
        lt.Cmd(strcat('"X Average Bin Size"=',num2str(actualHGS)));
        lt.Cmd(strcat('"Y Average Bin Size"=',num2str(actualVGS)));
        lt.Cmd('"Do Noise Reduction"=No ');
        lt.Cmd('\Q');
        lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].SPATIAL_LUM_METER[Spatial Lum Meter]"');
        lt.Cmd('"Meter Collection Mode"="Fixed Aperture" ');
        lt.Cmd('ApertureDefType="Distance and Radius" ');
        lt.Cmd(strcat('"Disk Radius"=',num2str(pupilSize),' '));
        lt.Cmd(strcat('Distance=',num2str(WDRLT),' '));
        lt.Cmd(strcat('Long=',num2str(HVALT),' ')); 
        lt.Cmd(strcat('Lat=',num2str(VVALT),' '));
        lt.Cmd('\Q');
    end
    cprintf('err',"����\n")
else % (���R��������)
    ReceiverCheck(lt)
end    
break
end
%% Begin Simulation
while (1)
cprintf('key',"LT ������......")
tStartSimulation = tic;
lt.Cmd("\V3D");
lt.Cmd("BeginAllSimulations");
tEndSimulation = toc(tStartSimulation);
cprintf('err',strcat("���� (",num2str(tEndSimulation)," seconds)\n"))
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
%% �}�ҦV�����-�Ŷ����� (optional)
% lt.Cmd("\V3D");
% lt.Cmd(['LumViewSpatialLuminanceChart "R13 �V��_����"']);
%% �s�����
while (1)
if saveLTRawData==1
    cprintf('key',"�g�J��Ƥ�......")
    dataArrayAll = cell(13,1);
    for N = NumReceiverArray
        meshKey = strcat("SURFACE_RECEIVER[R",num2str(Wr(N)),"].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance].SPATIAL_LUMINANCE_MESH[Spatial_Luminance_Mesh]");
        Xdim = lt.DbGet(meshKey, "X_Dimension");
        Ydim = lt.DbGet(meshKey, "Y_Dimension");
        imageFilepath = fullfile(fullDirc,strcat("R",num2str(Wr(N)),".png"));   % �s���G ��m
        xlsstr = strcat("Result_VVA",num2str(VVA),"_HVA",num2str(HVA),"_WDR",num2str(WDR),"_MRM",num2str(MRM),"_",dateString,".xlsx");
        excelFilepath = fullfile(fullDirc,xlsstr);  % �s���G ��m
        %% ��XExcel
        % �����
        dataArray=zeros(Ydim,Xdim);  %�ʺA�}�C
        [~,dataArray] = lt.GetMeshData(meshKey, dataArray(), "CellValue");
        dataArray=rot90(double(dataArray)); % ���n�B�z
        dataArrayAll{N} = dataArray;
        % �b��
        xAxisFirst = str2double(string(lt.DbGet(meshKey,"XCellCenterAt",def,1)));
        xAxisFinal = str2double(string(lt.DbGet(meshKey,"XCellCenterAt",def,Xdim)));
        xAxisArray = linspace(xAxisFirst,xAxisFinal,Xdim);
        yAxisFirst = str2double(string(lt.DbGet(meshKey,"YCellCenterAt",def,1)));
        yAxisFinal = str2double(string(lt.DbGet(meshKey,"YCellCenterAt",def,Ydim)));
        yAxisArray = linspace(yAxisFirst,yAxisFinal,Ydim)';
        strAtA1 = {'Y\X'};
        writecell(strAtA1,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','A1','UseExcel',true,'AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','B1','UseExcel',true,'AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(yAxisArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','A2','UseExcel',true,'AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(dataArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','B2','UseExcel',true,'AutoFitWidth',false); % �f�t LT �g�X���
        %% �����
        maxValue = max(max(dataArray));
        minValue = min(min(dataArray));
        LTImage = uint8((dataArray./maxValue)*255);               % 0��̧C��
        imwrite(LTImage,imageFilepath);
    end
    cprintf('err',"����\n")
end 
break
end
%% �O���p�Ȼ~�t (20220609 By GY)
while (1)
errorPeakPercentageArray = zeros(5,5);
actualRayHitArray = zeros(5,5);
actualRayMultiplierArray = zeros(5,5);

errorCheck = 0;
for N = NumReceiverArray
    ReceiverMeshKey = strcat("LENS_MANAGER[1].ILLUM_MANAGER[Illumination_Manager].RECEIVERS[Receiver_List].SURFACE_RECEIVER[R",num2str(Wr(N)),"].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance].SPATIAL_LUMINANCE_MESH[Spatial_Luminance_Mesh]");
    [row,column] = ind2sub([5,5],Wr(N));
    errorPeakPercentageArray(row,column) = lt.DbGet(ReceiverMeshKey,"ErrorAtPeak_Percent");
    ReceiverKey = strcat("LENS_MANAGER[1].ILLUM_MANAGER[Illumination_Manager].RECEIVERS[Receiver_List].SURFACE_RECEIVER[R",num2str(Wr(N)),"].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance]");
    actualRayHitArray(row,column) = lt.DbGet(ReceiverKey,"Actual_Ray_Hits");
    actualRayMultiplierArray(row,column) = lt.DbGet(ReceiverKey,"Actual_Ray_Multiplier");
    if errorCheck == 0 && actualRayHitArray(row,column) ~= rayNum
        errorCheck = 1;
        warning("���������S���F�ܳ]�w���u��")
    end
end

break
end
%% ��]�P���R
while (1)
cprintf('key',"��]�P���R......")
LTWDR = zeros(5,10);
intensityArray = zeros(2,13);
Wr = [1 5 7 8 9 12 13 14 17 18 19 21 25];
Datalist = zeros(5,13);
Datalist(1,:) = Wr;
F_each = cell(13,1);
maxIntensityArray = zeros(5,5);
for temptemp = 1:13 % ��l��
    F_each{temptemp} = zeros(Ydim,Xdim);
end
picTargetErodeArray = cell(5,5);

for k = NumReceiverArray

    temp = dataArrayAll{k};
    I = temp;
    Datalist(2,k)=roundn(max(temp(:)),-2);
    Datalist(3,k)=roundn(min(temp(:)),-2);
    Datalist(4,k)=roundn(min(temp(:))/max(temp(:)),-2);
    Datalist(5,k)=roundn((min(temp(:))+max(temp(:)))/(max(temp(:))-min(temp(:))),-2);
    if k == NumReceiverArray(1)
        Picc = size(temp);
        Pic = zeros(Picc(1)*5);
        picTargetErodeArray(:,:) = {uint8(zeros(Picc(1),Picc(2)))};
    end
    XX = uint8((temp./max(temp(:)))*255);
    Xmin= double(min(temp(:)))/double(max(temp(:)));

    F_each{k,1}=XX;

    %% Erosion
    if erosionCS == 1
        [row,column] = ind2sub([5,5],Wr(k));
        XX = imerode(XX,se);
        imageFilepath = fullfile(fullDirc,strcat("R",num2str(Wr(k)),"_Erosion.png"));   % �s���G ��m
        imwrite(XX,imageFilepath);
        picTargetErodeArray{row,column} = XX;
    end

    %% �ֿn��Ƨ�10% & 90%
    H_forobserve=histogram(XX);
    H=histogram(XX,'Normalization','cdf');
    edge=[H.BinLimits(1):0.5:H.BinLimits(end)];
    H=histogram(XX,'Normalization','cdf','BinEdges',edge);
    CumData_Y=H.Values;
    CumData_X=zeros(1,length(H.Values));
    for jj=1:length(H.BinEdges)-1
        CumData_X(jj)=0.5*(H.BinEdges(jj)+H.BinEdges(jj+1));
    end
    x_const=0:5:255;  %sim �p���ȥi��ݽվ� 0.1 �j���i�H5  !!!!!
    y_const=x_const;
    y_const(:)=0.1;
    Imin=polyxpoly(x_const,y_const,CumData_X,CumData_Y);
    y_const(:)=0.9;
    Imax=polyxpoly(x_const,y_const,CumData_X,CumData_Y);
    %%
    CR=(Imax-Imin)/(Imax+Imin);   %��X����
    CSF1=1/CR;
    LTWDR(Wr(k)) = roundn(CSF1,-1);
    intensityArray(1,k)=Imax;
    intensityArray(2,k)=Imin;
    
    W_non0 = LTWDR(LTWDR(:,1:5)>0);
    W_non0_AVG=round(sum(W_non0)/13,1);

    %% Imax Ratio 20220609 By GY
    [row,column] = ind2sub([5,5],Wr(k));
    maxIntensityArray(row,column) = Datalist(2,k);
    if k == NumReceiverArray(1) % �����Ĥ@�դ����I�j�� (��m13) (��7�I)
        maxIntensitySpecific = maxIntensityArray(row,column);
    end
end
if erosionCS == 1
    disp("CS Value (after erosion)");
elseif erosionCS == 0
    disp("CS Value");
end
disp(LTWDR);
disp(strcat("����: ",num2str(W_non0_AVG)));
maxIntensityArrayNormal2Each = maxIntensityArray/max(maxIntensityArray(:));
maxIntensityArrayNormal2Specific = maxIntensityArray/maxIntensitySpecific;

%% Summary
cprintf('key',"�g�JSummary��......")
F_zero=zeros(Picc(1),Picc(2));
Total_F=[F_each{1},F_zero,F_zero,F_zero,F_each{12};
        F_zero,F_each{3},F_each{6},F_each{9},F_zero;
        F_zero,F_each{4},F_each{7},F_each{10},F_zero;
        F_zero,F_each{5},F_each{8},F_each{11},F_zero;
        F_each{2},F_zero,F_zero,F_zero,F_each{13}];

figure;  imshow(Total_F,[0 255]);
imwrite(Total_F,fullfile(fullDirc,"13�I��.png"));

writematrix(LTWDR,excelFilepath,'Sheet','Summary!!','UseExcel',true,'AutoFitWidth',false);
writematrix(W_non0_AVG,excelFilepath,'Sheet','Summary!!','Range','A7','UseExcel',true,'AutoFitWidth',false);
txt={'Max';'Min';'Min/Max(%)';'CS (100%-0%)'};
writecell(txt,excelFilepath,'Sheet','Summary!!','Range','A12','UseExcel',true,'AutoFitWidth',false);
writematrix(Datalist,excelFilepath,'Sheet','Summary!!','Range','B11','UseExcel',true,'AutoFitWidth',false);
writecell({'�p�Ȼ~�t(%)'},excelFilepath,'Sheet','Summary!!','Range','P3','UseExcel',true,'AutoFitWidth',false);
writecell({'��ڱ������u��'},excelFilepath,'Sheet','Summary!!','Range','W3','UseExcel',true,'AutoFitWidth',false);
writecell({'�w���������u��'},excelFilepath,'Sheet','Summary!!','Range','W2','UseExcel',true,'AutoFitWidth',false);
writecell({'�̤j���զ���'},excelFilepath,'Sheet','Summary!!','Range','AD3','UseExcel',true,'AutoFitWidth',false);
writecell({'�̤j���j normal2self'},excelFilepath,'Sheet','Summary!!','Range','P10','UseExcel',true,'AutoFitWidth',false);
writecell({'�̤j���j normal2all'},excelFilepath,'Sheet','Summary!!','Range','W10','UseExcel',true,'AutoFitWidth',false);
writematrix(errorPeakPercentageArray,excelFilepath,'Sheet','Summary!!','Range','Q3','UseExcel',true,'AutoFitWidth',false);
writematrix(actualRayHitArray,excelFilepath,'Sheet','Summary!!','Range','X3','UseExcel',true,'AutoFitWidth',false);
writematrix(rayNum,excelFilepath,'Sheet','Summary!!','Range','X2','UseExcel',true,'AutoFitWidth',false);
writematrix(actualRayMultiplierArray,excelFilepath,'Sheet','Summary!!','Range','AE3','UseExcel',true,'AutoFitWidth',false);
writematrix(maxIntensityArrayNormal2Each,excelFilepath,'Sheet','Summary!!','Range','Q10','UseExcel',true,'AutoFitWidth',false);
writematrix(maxIntensityArrayNormal2Specific,excelFilepath,'Sheet','Summary!!','Range','X10','UseExcel',true,'AutoFitWidth',false);
cprintf('err',"����\n")
break
end
%% error peak ����
while (1)
if errorCheck == 1
    disp(strcat("max(�̤j���զ���) = ",num2str(max(actualRayMultiplierArray(:)))));
else
    disp("�Ҧ��������F�ܳ]�w���u��.")
end
disp("�p�Ȼ~�t")
disp(errorPeakPercentageArray)
break
end
%% ���� Imax Imin (13�I) (For CS<mean>)
intensityArrayEachSeed{whichSeedNum} = intensityArray;
CSArrayEachSeed{whichSeedNum} = LTWDR(:,1:5);

if erosionCS == 1
picTargetErodeCombine = cell2mat(picTargetErodeArray);
imageFilepath = fullfile(fullDirc,"combination_Erosion.png");   % �s���G ��m
imwrite(picTargetErodeCombine,imageFilepath);
end
tEndEachSeed = toc(tStartEachSeed);
disp(strcat("process complete (each seed): ",num2str(tEndEachSeed)," seconds."))
end % seed ����
%% CS <mean>
Wr = [1 5 7 8 9 12 13 14 17 18 19 21 25];
meanCSFromIArray = zeros(5,5);
maxCSFromIArray = zeros(5,5);
minCSFromIArray = zeros(5,5);
meanCSArray = zeros(5,5);

for k = NumReceiverArray
    [row,column] = ind2sub([5,5],Wr(k));
    ImaxPool = zeros(1,seedNum);
    IminPool = zeros(1,seedNum);
    CSPool = zeros(1,seedNum);
    for kk = 1:seedNum
        tempArray = intensityArrayEachSeed{kk};
        ImaxPool(kk)= tempArray(1,k);
        IminPool(kk)= tempArray(2,k);
        tempArray = CSArrayEachSeed{kk};
        CSPool(kk) = tempArray(row,column);
    end
    maxImax = max(ImaxPool);
    minImax = min(ImaxPool);
    maxImin = max(IminPool);
    minImin = min(IminPool);
    maxCSFromI = (minImax+maxImin)/(minImax-maxImin);
    minCSFromI = (maxImax+minImin)/(maxImax-minImin);
    meanCSFromI = 0.5*(maxCSFromI+minCSFromI);
    meanCSFromI = roundn(meanCSFromI,-1);
    meanCSFromIArray(row,column) = meanCSFromI;
    maxCSFromIArray(row,column) = maxCSFromI;
    minCSFromIArray(row,column) = minCSFromI;
    meanCSArray(row,column) = mean(CSPool);
end
CSMeanVAArray{whichVANum} = meanCSArray;
tEndVASet = toc(tStartVASet);
disp(strcat("process complete (each VASet): ",num2str(tEndVASet)," seconds."))
end % VA �� ����
cprintf('text',"�{�ǧ���\n")
disp(strcat("�`��O�ɶ�: ",num2str(toc(tStartFromBegining))," ��"))
%% Function
function [WDR,theta,phi] = Position2Angle(viewPointX,viewPointY,WDZ)
% Position2Angle: �������лP�y�y���ഫ (By GY 20220613)
% < Function Handle Example >
% P2A = @Position2Angle;
% [vd,theta,phi] = P2A(X,Y,Z);
% X,Y: X Y coordinate; (�k�⧤�Шt�Y�i)
% Z: Z coordinate where system top is set to z=0;
% (LT Version) (should be checked for other usage)
%%
arguments
    viewPointX {mustBeNumeric}
    viewPointY {mustBeNumeric}
    WDZ {mustBeNumeric}
end

%% 
% theta: 0 ~ 180
% phi: 0 ~ 360

[azimuth,elevation,WDR] = cart2sph(viewPointX,viewPointY,WDZ);
phi = rad2deg(azimuth);
theta = 90 - rad2deg(elevation); % elevation --> theta

    
end
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