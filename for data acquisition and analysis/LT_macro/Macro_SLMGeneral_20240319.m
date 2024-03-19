%% Original Author: BM %%
% 20220615: Modified by GY (project since 20220615)
% �γ~: �����ѪR20�I
% �γ~: LT�٭�v��
% �����ѪR20�I: II �ϩR�W�п�u: ... _WDR700_VVA40.5_HVA0_ ...
% Last Update: 20240319
% content: meshgrid fix (isequal(A,""))
close all;clear;clc
tStartFromBegining = tic;

%% �ϥΪ̿�J
LTID = 12840;
saveLTFile = 0;                                 % �s�� LT File �}��
saveLTRaw = 1;                                  % �s�� LT excel, Fig. �}��
% root folder: will create folder here
rootFolder = "";                                % ����Ƨ��M�w. "": choose folder; cd: current folder (���Τ޸�); "abc/def...": other path
    customLineFolder = '';                      % ���G��Ƨ��B�~�W��: "MRM5 LT�٭� Custom..."
    customLineFile = '';                        % ���G�ɮ��B�~�W��: "(LT�٭�) II�W Custom..."
customDirectionGridApodizer = 0;                % �������פ��G (0:Lambertian�A1:�ϥΪ̦ۭq�A-1:�����{�����߰�)
customSurfaceGridApodizer = 1;                  % �����Ŷ����G (0:���áA1:�ϥΪ̦ۭq�A-1:�����{�����߰�)
    % ���k�ɰѼ�
    TXTSwitch = 1;                              % ���k�ɶ}��,�ҥνп�J�U��T�C�Ѽ�
        Zero2One = 1;                           % �O�_�N Matrix ���� 0 �ର 1
        HorSize = 165.24;                       % (mm)
        VerSize = 293.76;                       % (mm)
seed = 1;                                       % ���l�ؤl�X�A�i�����N����� (default = 1)

% �������Ѽ� (�C�����槡�|�R������)
smoothFactor = 0;                               % �O�_���}����. 0: ����, other: 3,5,7,...21 (���_��)
humanFactor = 1;                                % �O�_�}�ҤH�] Mesh (�M�ΤH��������v 1/120 ��)
moduleTop = 7.77;                               % �t�γ̳� Z �y�� (mm) (�������|�ۦ� + 0.001 mm)
xOffset = 0;                                    % Receiver LT_X offset (Hor) (�w�] LT �����k����) (mm)
yOffset = 0;                                    % Receiver LT_Y offset (Ver) (�w�] LT �����W����) (mm)
receiverSizeHor = 165.24;                       % �����������j�p (mm)
receiverSizeVer = 293.76;                       % �����������j�p (mm)
horGridSize = 0.0765;                           % set to "" if horGridNum / verGridNum have value (mm)
verGridSize = 0.0765;                           % set to "" if horGridNum / verGridNum have value (mm)
horGridNum = "";                                % set to "" if horGridSize / verGridSize have value (#)
verGridNum = "";                                % set to "" if horGridSize / verGridSize have value (#)
MRM = 1;                                        % �̤j���զ���
expectedERR = 0.1;                              % �w���p�Ȼ~�t
rayFactor = 0.1;                                % ���u���B�~�[��

% �����Ѽ� (�V��Ž�)
eyeMode = 0;                                    % -1 0 1 �����k��
pupilSize = 2.5;                                % �b�| (mm)
IPD = 60;                                       % �Ⲵ���߶Z�� (mm)
extractVAFromII = 1;                            % �O�_�q�Ϥ��ɦW���^�� VA��T (���ɩR�W�ݥ]�t:EX... _WDR700_VVA40.5_HVA0_ ...)                             
    % if VAExtractFromII = 0 �H�U�ѼƤ~����:                                             
    WDR = 500;                                  % ��� ���߲� �� ���߭��O
    VVA = 30;                                   % ��� ���߲� �� ���߭��O
    HVA = 0;                                    % ��� ���߲� �� ���߭��O
    softwareVA = 0;                             % �O�_�M�� �n�� II VA �^���Ҧ�
systemTiltAngle = 0;                            % �N VA �ഫ�����Ĩ���

% �Ҧ����
% wedge prism �Ҧ� (�������������� WP �׭�)
% assume wp PRA = 0
wp_mode = 0;
    wp_ver_size = 293.76;
    wp_PBA = 5;

% ��v�Ҧ� (���������沴���¦V) (���ŦX����v��)
% �O�_�վ㱵�����Ϭݨ쪺�٭쵲�G�󱵪�u�걡��?
projectRec = 0;
    panelHor = 165.24;
    panelVer = 293.76;

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% ���u�ƹw�� ��L�ѼƽT�{
while (1)
dateString = datestr(now,'mm-dd-yyyy HH-MM');

smoothFactorPool = [0,3,5,7,9,11,13,15,17,19,21];
if ~any(smoothFactor==smoothFactorPool)
    error("���ưѼƳ]�w���~")
end

if ~isequal(customLineFile,"")
    customLineFile = strcat("_",customLineFile);
end

if wp_mode == 1 && projectRec == 1
    error("[���~] ���i�P�ɶ}�� wp_mode �M projectRec mode. (�t�ΰ���)")
end

break
end

%% �Ыظ�Ƨ�
while (1)
if ~isempty(customLineFolder)
    customLineFolder = strcat("_",customLineFolder);
end
targetDirc = strcat("LT�٭� ",customLineFolder);  %�۩w�q��Ƨ��W��
rootFolder = string(rootFolder);
if isempty(rootFolder) || isequal(rootFolder,"")
    rootFolder = uigetdir("","��ܥؼи�Ƨ� (�N�b�Ӹ�Ƨ����إ�LT�٭��Ƨ�)");
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

%% seed info 20231019
seed_firstTime = 1;
disp(strcat("[info]: ���l�ؤl�X: ",num2str(seed)))
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
    excelFilepath = fullfile(fullDirc,strcat("(LT�٭�) ",name,customLineFile,".xlsx"));  % �s���G ��m
    checkExcelFileName(excelFilepath);
    imageFilepath = fullfile(fullDirc,strcat("(LT�٭�) ",name,customLineFile,".png"));   % �s���G ��m
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
    
    %% �ѼƧ�s (�ǳƫإ߱�����)
    % Update Term 1: WDR VVA HVA %
    while (1)
    if extractVAFromII == 1 && SGType == "Uniform"
        error("extractVAFromII cannot be 1 when the suface apodizer is 'uniform' mode (�t�ΰ���)")
    elseif extractVAFromII == 1 && SGType == "Grid"
        if softwareVA == 0
            % �T�O �R�W�]�t: ... _WDR700_VVA40.5_HVA0_ ...
            WDRPattern = caseInsensitivePattern("_WDR");
            VVAPattern = caseInsensitivePattern("_VVA");
            HVAPattern = caseInsensitivePattern("_HVA");
            WDRString = extractBetween(name,WDRPattern,VVAPattern);
            VVAString = extractBetween(name,VVAPattern,HVAPattern);
            HVAString = extractBetween(name,HVAPattern,"_");
            HVA = str2double(HVAString);
            if isempty(HVA) % �䴩 _WDR700_VVA40.5_HVA0.png �Φ� (HVA��S���Ѽ�)
                HVAString = extractAfter(name,HVAPattern);
            end
            % value check
            WDR = str2double(WDRString); 
            if isempty(WDR);beep;error("cannot extract WDR value. (�T�O�R�W�]�t:... _WDR700_VVA40.5_HVA0_ ...)");end
            VVA = str2double(VVAString); 
            if isempty(VVA);beep;error("cannot extract VVA value.  (�T�O�R�W�]�t:... _WDR700_VVA40.5_HVA0_ ...)");end
            HVA = str2double(HVAString); 
            if isempty(HVA);beep;error("cannot extract HVA value.  (�T�O�R�W�]�t:... _WDR700_VVA40.5_HVA0_ ...)");end
        elseif softwareVA == 1
            % �T�O �R�W�]�t: ... _VD=0700_VVA=40.5_HVA=0_ ...
            WDRPattern = caseInsensitivePattern("_VD=");
            VVAPattern = caseInsensitivePattern("_VVA=");
            HVAPattern = caseInsensitivePattern("_HVA=");
            WDRString = extractBetween(name,WDRPattern,VVAPattern);
            VVAString = extractBetween(name,VVAPattern,HVAPattern);
            HVAString = extractBetween(name,HVAPattern,"_");
            HVA = str2double(HVAString);
            if isempty(HVA) % �䴩 _HVA0.png �Φ� (HVA ��S���Ѽ�)
                HVAString = extractAfter(name,HVAPattern);
            end
            % value check
            WDR = str2double(WDRString); 
            if isempty(WDR);beep;error("cannot extract WDR value. (�T�O�R�W�]�t:... _VD=0700_VVA=40.5_HVA=0_ ...)");end
            VVA = str2double(VVAString); 
            if isempty(VVA);beep;error("cannot extract VVA value.  (�T�O�R�W�]�t:... _VD=0700_VVA=40.5_HVA=0_ ...)");end
            HVA = str2double(HVAString); 
            if isempty(HVA);beep;error("cannot extract HVA value.  (�T�O�R�W�]�t:... _VD=0700_VVA=40.5_HVA=0_ ...)");end
        end
    elseif extractVAFromII == 0
        WDR;VVA;HVA; % �ϥΪ̪�����J
    end
    % �Ҽ{ tilt
    if (N == 1 && extractVAFromII == 0) || extractVAFromII == 1 % 20230323
        VVA_Ori = VVA; % ������l���ץ�
        HVA_Ori = HVA; % ������l���ץ�
        [VVA,HVA] = TiltAngle(WDR,VVA,HVA,systemTiltAngle);
    end
    break
    end
    
    % Update Term 2: �����Ҧ� �P ���O Offset %
    while (1)
    receiverZ = moduleTop + 0.001;
    viewPointX = WDR*sind(VVA)*cosd(HVA); % ������VVA
    viewPointY = WDR*sind(VVA)*sind(HVA); % ������VVA
    pupilCenterXY = [0;eyeMode*IPD*0.5]; %�H���y���߬��s�I
    % �Ⲵ����
    % ���������: no Prism ��HVA�� (HVA=90 && VVA ~=0)
    if HVA ~= 90
        pupilCenterXYZTemp=[pupilCenterXY;0];
        pupilCenterXYZRotTemp = rotz(HVA) * pupilCenterXYZTemp;
        pupilCenterXY = pupilCenterXYZRotTemp(1:2); % �����V�q
    else
        pupilCenterXY = [0;0];
    end
    % LT ���Шt
    eyePosition = [WDR*sind(VVA)*sind(HVA),-WDR*sind(VVA)*cosd(HVA),receiverZ+WDR*cosd(VVA)]...
        + [pupilCenterXY(2),-pupilCenterXY(1),0];
    receiverPosition = [xOffset,yOffset,receiverZ];
    incidentVector = eyePosition - receiverPosition;
    P2A = @Position2Angle; % ML ���Шt
    [wdrEach,thetaEach,phiEach] = P2A(-incidentVector(2),incidentVector(1),incidentVector(3));
    WDRLT = wdrEach; VVALT = thetaEach; HVALT = phiEach;
    HVALT = mod(360-HVALT,360);    
    % ��ڸg�n�שM�u�@�Z�� (���߲��襤�߭��O): WDR, VVA, HVA
    % LT�nKeyin���g�n�שM�u�@�Z��: WDRLT, VVALT, HVALT
    break
    end
    
    % Update Term 3: Mesh ���u�� %
    while (1)
    if humanFactor == 1
        angRes = tand(1/120); % �H��������v
        HFGridSize = WDRLT * angRes * 2;
        actualHGS = HFGridSize;
        actualVGS = HFGridSize;
        actualHGN = round(receiverSizeHor/actualHGS);
        actualVGN = round(receiverSizeVer/actualVGS);
    elseif humanFactor == 0
        if isequal(horGridNum,"") && isequal(verGridNum,"")
            actualHGS = horGridSize;
            actualVGS = verGridSize;
            actualHGN = round(receiverSizeHor/horGridSize);
            actualVGN = round(receiverSizeVer/verGridSize);
        elseif isequal(horGridSize,"") && isequal(verGridSize,"")
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
    
    % update Term 4: ��v�Ҧ� Receiver �Ѽƭp�� % By Louie
    while projectRec == 1 
        WDX = WDRLT*abs(sind(VVALT)*sind(HVALT));
        WDY = WDRLT*abs(sind(VVALT)*cosd(HVALT));
        WDZ = WDRLT*abs(cosd(VVALT));
        dummyAlpha = atand(WDY/WDZ);
        dummyBeta = asind(WDX/WDRLT);
        shift_Z = sind(dummyAlpha)*panelVer/2+sind(dummyBeta)*panelHor/2*cosd(dummyAlpha);
        shift_X = shift_Z/WDZ*WDX;
        shift_Y = shift_Z/WDZ*WDY;

        WDRLT = (WDZ-shift_Z)/abs(cosd(VVALT));
        xOffset = shift_X;
        yOffset = -shift_Y;
        zOffset = shift_Z+moduleTop;
        dummyAlpha = -dummyAlpha;
        dummyBeta = -dummyBeta;
        dummyGamma = HVALT;
        if HVA==0 && VVA==0
            dummyGamma=dummyGamma+90;
        end
        break
    end
    
    % update Term 5: WP �Ҧ�
    while wp_mode == 1
        % 1. �������W�� wedge height (dummy_z)
        % 2. �������b���� wp_PBA (dummy_alpha)
        % ���ⵥ�Ĳ����y��
        
        % �p�Ⱶ������m�� Z (�� [xOffset,yOffset,receiverZ] �M�w)
        % yOffset = 0 wp_height = wp �b��, y �V�W����
        % y > 0 wp_h < wp �b��
        we_height = (0.5 * wp_ver_size - yOffset) * tand(wp_PBA);

        % �p�ⵥ�Ĳ���
        E0 = [eyePosition(1:2),WDR*cosd(VVA)];
        E1 = E0 - [xOffset,yOffset,we_height];
        E2 = (rotx(-wp_PBA)^(-1) * E1')'; % ���⧤�Шt wp_PBA --> -wp_PBA
        
        % ���� VA
        [wdr_wp,theta_wp,phi_wp] = P2A(-E2(2),E2(1),E2(3));
        WDRLT = wdr_wp; VVALT = theta_wp; HVALT = phi_wp;
        HVALT = mod(360-HVALT,360); 

        % dummy plane �Ѽ� assume "wp_PRA = 0"
        zOffset = moduleTop + 0.001 + we_height;
        dummyAlpha = wp_PBA;
        dummyBeta = 0;
        dummyGamma = 0;
        break
    end
    
    %% �������إ� %% For 20�I���R: �C��Loop WDR���@�� ���n��s
    while (1)
    cprintf('key',"�إ߱�����......")
    % Dummy Plane: "D"; Receiver: "R"
    ReceiverDelete(lt); % �R����e������
    lt.Cmd("\V3D"); % ����@�ɮy�Ф���
    % dummy plane �إ�
    lt.Cmd('DummyPlane ');
    lt.Cmd('XYZ'); % ���������I
    lt.Cmd(strcat(num2str(xOffset),',',num2str(yOffset),',',num2str(moduleTop+0.001)));
    lt.Cmd('XYZ'); % �������k�u�I (�P���I�۴�k�u)
    lt.Cmd(strcat(num2str(xOffset),',',num2str(yOffset),',',num2str(moduleTop+0.001+1)));
    lt.Cmd('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"');
    lt.Cmd('Name=D');
    lt.Cmd('\Q');
    lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Width",receiverSizeHor);
    lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Height",receiverSizeVer);
    if projectRec == 1 || wp_mode == 1
        lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Z",zOffset);
        lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Alpha",dummyAlpha);
        lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Beta",dummyBeta);
        lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Gamma",dummyGamma);
    end
    % ������ �إ�
    lt.Cmd(strcat('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"'));
    lt.Cmd('"Add Receiver"=');
    lt.Cmd('\Q');
    lt.Cmd('\O"LENS_MANAGER[1].ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last]"');    %Follow 2*N
    lt.Cmd('Name=R');
    lt.Cmd('Responsivity=Photometric '); % ��� ���q�q
    %lt.Cmd('Responsivity=Radiometric '); % ��� ��g�q�q
%     lt.Cmd('"Photometry Type"="Photometry Type B" '); % ��V:��������
    lt.Cmd('\Q');
    % �����V�e����
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
    if smoothFactor == 0
        lt.Cmd('"Do Noise Reduction"=No ');
    else
        lt.Cmd('"Do Noise Reduction"=Yes ');
        lt.Cmd(strcat('"Kernel Size N"=',num2str(smoothFactor),' '));
    end
    lt.Cmd('\Q');
    lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].SPATIAL_LUM_METER[Spatial Lum Meter]"');
    lt.Cmd('"Meter Collection Mode"="Fixed Aperture" ');
    lt.Cmd('ApertureDefType="Distance and Radius" ');
    lt.Cmd(strcat('"Disk Radius"=',num2str(pupilSize),' '));
    lt.Cmd(strcat('Distance=',num2str(WDRLT),' ')); % update each loop
    if projectRec == 0
        lt.Cmd(strcat('Long=',num2str(HVALT),' '));     % update each loop
        lt.Cmd(strcat('Lat=',num2str(VVALT),' '));      % update each loop
    elseif projectRec == 1 % �����k�V�q�P����������
        lt.Cmd(strcat('Long=',"0",' '));     % update each loop
        lt.Cmd(strcat('Lat=',"0",' '));      % update each loop
    end
    lt.Cmd('\Q');
    %% 20231019 seed �]�w
    while (seed_firstTime == 1)
        lt.Cmd("\V3D"); % ����@�ɮy�Ф���
        lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].SIMULATIONS[ForwardAll]" ');
        lt.Cmd(strcat("StartingSeed=",num2str(seed)));
        lt.Cmd('\Q');
        seed_firstTime = 0;
    break
    end
    cprintf('err',"����\n")
    break
    end
    disp(strcat("��e�B�z:"," WDR",num2str(WDR)," VVA",num2str(VVA_Ori)," HVA",num2str(HVA_Ori)," STA",num2str(systemTiltAngle)))
    if systemTiltAngle ~= 0 
        disp(strcat(" (���Ĩ���: VVA",num2str(VVA)," HVA",num2str(HVA),")"))
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
    errActual = lt.DbGet("SURFACE_RECEIVER[R].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance].SPATIAL_LUMINANCE_MESH[Spatial Luminance Mesh]","ErrorAtPeak_Percent");
    cprintf('text',"�p�Ȼ~�t: %.2f %% \n",round(errActual*100)/100)
    ARM = lt.DbGet("SURFACE_RECEIVER[R].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance]","Actual_Ray_Multiplier");
    if MRM < ARM;warning("�̤j���u���զ��� �p�� ��ڥ��u���զ���");end
    %% �s��LT �ɮ�
    while (1)
    if saveLTFile==1
        cprintf('key',"[info] LT �s�ɤ�......")
        backupFileName = strcat("LTFile Fig",num2str(N)," ",dateString," (Backup)");
        lt.SetOption('ShowFileDialogBox', 0);     % �۰ʦs�� ���s�X�s�ɵ��� (���n)
        lt.Cmd('\VConsole');
        lt.Cmd(strcat("SaveAs """,fullfile(fullDirc,backupFileName),""""));
        lt.SetOption('ShowFileDialogBox', 1);     % �^�_�]�w (���n)
        cprintf('err',"����\n")
        dips(strcat("[info] �ɦW: ",backupFileName,".lts"));
    end
    break
    end
    %% �s���V�����Raw Datas & LTImage Photos
    while (1)
    if saveLTRaw==1
        cprintf('key',"LT Raw Data �s�ɤ�......")
        meshkey = "SURFACE_RECEIVER[R].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance].SPATIAL_LUMINANCE_MESH[Spatial_Luminance_Mesh]";
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
        if isfile(excelFilepath);delete(excelFilepath);end
        writecell(strAtA1,excelFilepath,'Range','A1','AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Range','B1','AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(yAxisArray,excelFilepath,'Range','A2','AutoFitWidth',false); % �f�t LT �g�X���
        writematrix(DataArray,excelFilepath,'Range','B2','AutoFitWidth',false); % �f�t xlsx2png coding �榡
        cprintf('err',"����\n")
        %% �����
        cprintf('key',"�ץX���ɤ�......")
        maxValue = max(max(DataArray));
        minValue = min(min(DataArray));
        LTImage = uint8((DataArray./maxValue)*255);               % 0��̧C��
        imwrite(LTImage,imageFilepath);
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
function sourceName = LightSourceSetup(lt) % ������e�Ҧ�����
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
    lt.ListDelete(ObjectList);
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