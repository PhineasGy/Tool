%% Original Author: BM %%
% 20220615: Modified by GY (project since 20220615)
% 用途: 極限解析20點
% 用途: LT還原影像
% 極限解析20點: II 圖命名請遵守: ... _WDR700_VVA40.5_HVA0_ ...
% Last Update: 20240319
% content: meshgrid fix (isequal(A,""))
close all;clear;clc
tStartFromBegining = tic;

%% 使用者輸入
LTID = 12840;
saveLTFile = 0;                                 % 存取 LT File 開關
saveLTRaw = 1;                                  % 存取 LT excel, Fig. 開關
% root folder: will create folder here
rootFolder = "";                                % 母資料夾決定. "": choose folder; cd: current folder (不用引號); "abc/def...": other path
    customLineFolder = '';                      % 結果資料夾額外名稱: "MRM5 LT還原 Custom..."
    customLineFile = '';                        % 結果檔案額外名稱: "(LT還原) II名 Custom..."
customDirectionGridApodizer = 0;                % 光源角度分佈 (0:Lambertian，1:使用者自訂，-1:維持現狀不詢問)
customSurfaceGridApodizer = 1;                  % 光源空間分佈 (0:均勻，1:使用者自訂，-1:維持現狀不詢問)
    % 切趾檔參數
    TXTSwitch = 1;                              % 切趾檔開關,啟用請輸入下方三列參數
        Zero2One = 1;                           % 是否將 Matrix 中的 0 轉為 1
        HorSize = 165.24;                       % (mm)
        VerSize = 293.76;                       % (mm)
seed = 1;                                       % 光追種子碼，可為任意正整數 (default = 1)

% 接收器參數 (每次執行均會刪掉重建)
smoothFactor = 0;                               % 是否打開平滑. 0: 關閉, other: 3,5,7,...21 (限奇數)
humanFactor = 1;                                % 是否開啟人因 Mesh (套用人眼角分辨率 1/120 度)
moduleTop = 7.77;                               % 系統最頂 Z 座標 (mm) (接收面會自行 + 0.001 mm)
xOffset = 0;                                    % Receiver LT_X offset (Hor) (預設 LT 中往右為正) (mm)
yOffset = 0;                                    % Receiver LT_Y offset (Ver) (預設 LT 中往上為正) (mm)
receiverSizeHor = 165.24;                       % 接收面水平大小 (mm)
receiverSizeVer = 293.76;                       % 接收面垂直大小 (mm)
horGridSize = 0.0765;                           % set to "" if horGridNum / verGridNum have value (mm)
verGridSize = 0.0765;                           % set to "" if horGridNum / verGridNum have value (mm)
horGridNum = "";                                % set to "" if horGridSize / verGridSize have value (#)
verGridNum = "";                                % set to "" if horGridSize / verGridSize have value (#)
MRM = 1;                                        % 最大嘗試次數
expectedERR = 0.1;                              % 預期峰值誤差
rayFactor = 0.1;                                % 光線數額外加乘

% 眼睛參數 (向後空輝)
eyeMode = 0;                                    % -1 0 1 左中右眼
pupilSize = 2.5;                                % 半徑 (mm)
IPD = 60;                                       % 兩眼中心距離 (mm)
extractVAFromII = 1;                            % 是否從圖片檔名中擷取 VA資訊 (圖檔命名需包含:EX... _WDR700_VVA40.5_HVA0_ ...)                             
    % if VAExtractFromII = 0 以下參數才有效:                                             
    WDR = 500;                                  % 基於 中心眼 對 中心面板
    VVA = 30;                                   % 基於 中心眼 對 中心面板
    HVA = 0;                                    % 基於 中心眼 對 中心面板
    softwareVA = 0;                             % 是否套用 軟體 II VA 擷取模式
systemTiltAngle = 0;                            % 將 VA 轉換為等效角度

% 模式選擇
% wedge prism 模式 (接收面必須平行 WP 斜面)
% assume wp PRA = 0
wp_mode = 0;
    wp_ver_size = 293.76;
    wp_PBA = 5;

% 投影模式 (接收面平行眼睛朝向) (較符合實驗影像)
% 是否調整接收面使看到的還原結果更接近真實情形?
projectRec = 0;
    panelHor = 165.24;
    panelVer = 293.76;

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% 光線數預估 其他參數確認
while (1)
dateString = datestr(now,'mm-dd-yyyy HH-MM');

smoothFactorPool = [0,3,5,7,9,11,13,15,17,19,21];
if ~any(smoothFactor==smoothFactorPool)
    error("平滑參數設定有誤")
end

if ~isequal(customLineFile,"")
    customLineFile = strcat("_",customLineFile);
end

if wp_mode == 1 && projectRec == 1
    error("[錯誤] 不可同時開啟 wp_mode 和 projectRec mode. (系統停止)")
end

break
end

%% 創建資料夾
while (1)
if ~isempty(customLineFolder)
    customLineFolder = strcat("_",customLineFolder);
end
targetDirc = strcat("LT還原 ",customLineFolder);  %自定義資料夾名稱
rootFolder = string(rootFolder);
if isempty(rootFolder) || isequal(rootFolder,"")
    rootFolder = uigetdir("","選擇目標資料夾 (將在該資料夾中建立LT還原資料夾)");
    if rootFolder == 0;disp("系統停止");return;end
elseif ~isstring(rootFolder)
    beep
    error("'rootFolder' variable should be either string or char type. (系統停止)")
end
lastwarn('');
fullDirc = fullfile(rootFolder,targetDirc);
mkdir(fullDirc);
msg = lastwarn;
if isequal('Directory already exists.',msg)
    disp("< 自定義資料夾已存在當前目錄 >")
end
break
end

%% 連結Light Tool (Citrix version)
while (1)
def=System.Reflection.Missing.Value;
% ltcom64path=['C:\Program Files\Optical Research Associates\LightTools 9.1.1\Utilities.NET\LTCOM64.dll'];      %LTCOM64.dll路徑
ltcom64path=['C:\Program Files (x86)\Common Files\Optical Research Associates\LightTools\LTCOM64.dll'];      %LTCOM64.dll路徑_Critrix!!
asm=NET.addAssembly(ltcom64path);
lt=LTCOM64.LTAPIx;
lt.LTPID=LTID;
lt.UpdateLTPointer;
lt.Message('已成功連結LT (From Matlab)');
% 檢查 status
% [(val),...,status] = lt command
% check whether string(status) == "ltStatusSuccess" or not
% or check status == 0 or not (0 == success)
break
end

%% 讀取檔案: 角度分佈, 空間分佈模式決定
while (1)
% 面板光源名稱擷取
sourceName = LightSourceSetup(lt); % 取得當前光源名稱 (cube)
% 在光源"右面"切換切趾檔
source = strcat("CUBE_SURFACE_SOURCE[",sourceName,"]");
currentSG_first = string(lt.DbGet(strcat(source,".NATIVE_EMITTER[RightSurface]"),"Surface Apodizer Type"));
currentDG_first = string(lt.DbGet(strcat(source,".NATIVE_EMITTER[RightSurface]"),"Direction Apodizer Type"));
lengthFiles = 1;

%% 光源角度分佈設定
if customDirectionGridApodizer == 1
    cprintf("[info]: 光源角度分佈類型為: 使用者自訂 ...... ")
    DGType = "Grid";
    % 角度分佈切趾檔
    pathname_angle=[];
    [filename_angle, pathname_angle] = uigetfile(strcat(pathname_angle,'*.txt'), '請選擇光源角度分佈切趾檔');
    if ~ischar(pathname_angle);disp("系統停止");return;end
    filepath_angle = fullfile(pathname_angle, filename_angle);        % 圖檔/切趾檔 位置
elseif customDirectionGridApodizer == 0 
    disp("[info]: 光源角度分佈類型為: Lambertian")
    DGType = "Lambertian";
elseif customDirectionGridApodizer == -1 % 不更動當前設定
    str_temp1 = ["使用者自訂","Lambertian"];
    if currentDG_first == "Grid"
        str_temp1 = "使用者自訂";
        DGType = "Grid";
    elseif currentDG_first == "Lambertian"
        str_temp1 = "Lambertian";
        DGType = "Lambertian";
    end
    disp(strcat("[info]: 光源角度分佈類型為: ",str_temp1," (不更動)"))
else
    error("[error]: 無法識別光源角度分佈類型 (customDirectionGridApodizer) (系統停止)")
end
if customDirectionGridApodizer ~= -1
    lt.Cmd(strcat("\O",source,".NATIVE_EMITTER[RightSurface]")); 
    lt.Cmd(strcat('"Direction Apodizer Type"=',DGType,' ')); 
    if DGType == "Grid" % 匯入切趾檔
        lt.Cmd("\VConsole "); % 切到資訊介面
        lt.Cmd(strcat("\O",source,".NATIVE_EMITTER[RightSurface]"));
        lt.Cmd('"Direction Apodizer Type"=Grid ');
        lt.DbSet(strcat(source,".NATIVE_EMITTER[RightSurface].DIRECTION_GRID_APODIZER[DirectionGridApodizer]"), "LoadFileName", filepath_angle);
        % 檢查切趾檔有效性
        if lt.GetLastMsg(1) == "錯誤: 匯入網格失敗。"
            error("[error]: 無法識別光源角度分佈切趾檔 (系統停止)")
        end
        disp("匯入切趾檔成功!")
        lt.Cmd("\V3D "); % 切到世界座標介面
    end
end
%% 光源空間分佈設定
if customSurfaceGridApodizer == 1
    disp("[info]: 光源空間分佈類型為: 使用者自訂")
    SGType = "Grid";
    if TXTSwitch==1 % 抓II圖 (自動切趾)
        pathname=[];
        [filename, pathname] = uigetfile(strcat(pathname,'*.png'), '請選擇 II 圖作為光源空間分佈切趾檔 (可多選)','MultiSelect','on');
    elseif TXTSwitch==0 % 抓II切趾檔
        pathname=[]; 
        [filename, pathname] = uigetfile(strcat(pathname,'*.txt'), '請選擇光源空間分佈切趾檔 (可多選)','MultiSelect','on');
    end
    % 計算檔案數量
    if ~ischar(pathname) 
        disp("系統停止");return;end
    if ischar(filename)
        lengthFiles = 1;
    else
        lengthFiles = length(filename);
    end
elseif customSurfaceGridApodizer == 0
    disp("[info]: 光源空間分佈類型為: 均勻")
    SGType = "Uniform";
elseif customSurfaceGridApodizer == -1 % 不更動當前設定
    str_temp2 = ["使用者自訂","均勻"];
    if currentSG_first == "Grid"
        str_temp2 = "使用者自訂";
        SGType = "Grid";
    elseif currentSG_first == "Uniform"
        str_temp2 = "均勻";
        SGType = "Uniform";
    end
    disp(strcat("[info]: 光源空間分佈類型為: ",str_temp2," (不更動)"))
else
    error("[error]: 無法識別光源空間分佈類型 (customSurfaceGridApodizer) (系統停止)")
end
if  customSurfaceGridApodizer ~= -1
    lt.Cmd(strcat("\O",source,".NATIVE_EMITTER[RightSurface]"));
    lt.Cmd(strcat('"Surface Apodizer Type"=',SGType,' '));
end
break
end

%% seed info 20231019
seed_firstTime = 1;
disp(strcat("[info]: 光追種子碼: ",num2str(seed)))
%% Each File 處理 (包含模擬)
for N = 1:lengthFiles
    tStartEachLoop = tic;
    cprintf("-------------------------------\n")
    %% FullFile
    while (1)
    name = "";
    if customSurfaceGridApodizer == 1
        if lengthFiles == 1
            name = erase(filename,[".png",".bmp",".txt"]);
            filepath = fullfile(pathname, filename);        % 圖檔/切趾檔 位置
        else
            name = erase(filename{N},[".png",".bmp",".txt"]);
            filepath = fullfile(pathname, filename{N});     % 圖檔/切趾檔 位置
        end
    end
    excelFilepath = fullfile(fullDirc,strcat("(LT還原) ",name,customLineFile,".xlsx"));  % 存結果 位置
    checkExcelFileName(excelFilepath);
    imageFilepath = fullfile(fullDirc,strcat("(LT還原) ",name,customLineFile,".png"));   % 存結果 位置
    break
    end
    
    %% 自動轉成切趾檔 (if 輸入為圖檔)
    while (1)
    if SGType == "Grid" % customSurfaceGridApodizer 只可能是 1 or -1
        if customSurfaceGridApodizer == 1
            if TXTSwitch==1 % 圖檔 為輸入
                II = imread(filepath); % valuse: 0-255  
                cprintf('key',"影像轉成切趾檔中......")
                func_spa(II,filepath,HorSize,VerSize,Zero2One); % 寫出txt檔
                filepath = strcat(erase(filepath,[".png",".bmp"]),".txt");
            end
            % 匯入切趾檔
            lt.DbSet(strcat(source,".NATIVE_EMITTER[RightSurface].SURFACE_GRID_APODIZER[SurfaceGridApodizer]"), "LoadFileName", filepath);
            % 檢查切趾檔有效性
            if lt.GetLastMsg(1) == "錯誤: 匯入網格失敗。"
                disp("[error]: 錯誤的切趾檔")
                disp(filepath);
                error("[error]: 無法識別光源空間分佈切趾檔 (系統停止)")
            end 
            disp("[info] 空間切趾擋匯入成功!")
        elseif customSurfaceGridApodizer == -1 % 不做切趾檔
        end
        % 紀錄切趾檔名
        name = string(lt.DbGet(strcat(source,".NATIVE_EMITTER[RightSurface].SURFACE_GRID_APODIZER[SurfaceGridApodizer]"), "LoadFileName"));
        name = erase(name,".txt");
    end
    break
    end
    
    %% 參數更新 (準備建立接收器)
    % Update Term 1: WDR VVA HVA %
    while (1)
    if extractVAFromII == 1 && SGType == "Uniform"
        error("extractVAFromII cannot be 1 when the suface apodizer is 'uniform' mode (系統停止)")
    elseif extractVAFromII == 1 && SGType == "Grid"
        if softwareVA == 0
            % 確保 命名包含: ... _WDR700_VVA40.5_HVA0_ ...
            WDRPattern = caseInsensitivePattern("_WDR");
            VVAPattern = caseInsensitivePattern("_VVA");
            HVAPattern = caseInsensitivePattern("_HVA");
            WDRString = extractBetween(name,WDRPattern,VVAPattern);
            VVAString = extractBetween(name,VVAPattern,HVAPattern);
            HVAString = extractBetween(name,HVAPattern,"_");
            HVA = str2double(HVAString);
            if isempty(HVA) % 支援 _WDR700_VVA40.5_HVA0.png 形式 (HVA後沒有參數)
                HVAString = extractAfter(name,HVAPattern);
            end
            % value check
            WDR = str2double(WDRString); 
            if isempty(WDR);beep;error("cannot extract WDR value. (確保命名包含:... _WDR700_VVA40.5_HVA0_ ...)");end
            VVA = str2double(VVAString); 
            if isempty(VVA);beep;error("cannot extract VVA value.  (確保命名包含:... _WDR700_VVA40.5_HVA0_ ...)");end
            HVA = str2double(HVAString); 
            if isempty(HVA);beep;error("cannot extract HVA value.  (確保命名包含:... _WDR700_VVA40.5_HVA0_ ...)");end
        elseif softwareVA == 1
            % 確保 命名包含: ... _VD=0700_VVA=40.5_HVA=0_ ...
            WDRPattern = caseInsensitivePattern("_VD=");
            VVAPattern = caseInsensitivePattern("_VVA=");
            HVAPattern = caseInsensitivePattern("_HVA=");
            WDRString = extractBetween(name,WDRPattern,VVAPattern);
            VVAString = extractBetween(name,VVAPattern,HVAPattern);
            HVAString = extractBetween(name,HVAPattern,"_");
            HVA = str2double(HVAString);
            if isempty(HVA) % 支援 _HVA0.png 形式 (HVA 後沒有參數)
                HVAString = extractAfter(name,HVAPattern);
            end
            % value check
            WDR = str2double(WDRString); 
            if isempty(WDR);beep;error("cannot extract WDR value. (確保命名包含:... _VD=0700_VVA=40.5_HVA=0_ ...)");end
            VVA = str2double(VVAString); 
            if isempty(VVA);beep;error("cannot extract VVA value.  (確保命名包含:... _VD=0700_VVA=40.5_HVA=0_ ...)");end
            HVA = str2double(HVAString); 
            if isempty(HVA);beep;error("cannot extract HVA value.  (確保命名包含:... _VD=0700_VVA=40.5_HVA=0_ ...)");end
        end
    elseif extractVAFromII == 0
        WDR;VVA;HVA; % 使用者直接輸入
    end
    % 考慮 tilt
    if (N == 1 && extractVAFromII == 0) || extractVAFromII == 1 % 20230323
        VVA_Ori = VVA; % 紀錄初始角度用
        HVA_Ori = HVA; % 紀錄初始角度用
        [VVA,HVA] = TiltAngle(WDR,VVA,HVA,systemTiltAngle);
    end
    break
    end
    
    % Update Term 2: 眼睛模式 與 面板 Offset %
    while (1)
    receiverZ = moduleTop + 0.001;
    viewPointX = WDR*sind(VVA)*cosd(HVA); % 垂直方向VA
    viewPointY = WDR*sind(VVA)*sind(HVA); % 水平方向VA
    pupilCenterXY = [0;eyeMode*IPD*0.5]; %以眼球中心為零點
    % 兩眼旋轉
    % 不旋轉條件: no Prism 掃HVA時 (HVA=90 && VVA ~=0)
    if HVA ~= 90
        pupilCenterXYZTemp=[pupilCenterXY;0];
        pupilCenterXYZRotTemp = rotz(HVA) * pupilCenterXYZTemp;
        pupilCenterXY = pupilCenterXYZRotTemp(1:2); % 旋轉後向量
    else
        pupilCenterXY = [0;0];
    end
    % LT 坐標系
    eyePosition = [WDR*sind(VVA)*sind(HVA),-WDR*sind(VVA)*cosd(HVA),receiverZ+WDR*cosd(VVA)]...
        + [pupilCenterXY(2),-pupilCenterXY(1),0];
    receiverPosition = [xOffset,yOffset,receiverZ];
    incidentVector = eyePosition - receiverPosition;
    P2A = @Position2Angle; % ML 坐標系
    [wdrEach,thetaEach,phiEach] = P2A(-incidentVector(2),incidentVector(1),incidentVector(3));
    WDRLT = wdrEach; VVALT = thetaEach; HVALT = phiEach;
    HVALT = mod(360-HVALT,360);    
    % 實際經緯度和工作距離 (中心眼對中心面板): WDR, VVA, HVA
    % LT要Keyin的經緯度和工作距離: WDRLT, VVALT, HVALT
    break
    end
    
    % Update Term 3: Mesh 光線數 %
    while (1)
    if humanFactor == 1
        angRes = tand(1/120); % 人眼角分辨率
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
                warning("Mesh 數設定可能有誤: 目前以 horGridSize / verGridSize 為主推算 horGridNum / verGridNum")
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
    
    % update Term 4: 投影模式 Receiver 參數計算 % By Louie
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
    
    % update Term 5: WP 模式
    while wp_mode == 1
        % 1. 接收面上抬 wedge height (dummy_z)
        % 2. 接收面軸旋轉 wp_PBA (dummy_alpha)
        % 換算等效眼睛座標
        
        % 計算接收面位置的 Z (由 [xOffset,yOffset,receiverZ] 決定)
        % yOffset = 0 wp_height = wp 半高, y 向上為正
        % y > 0 wp_h < wp 半高
        we_height = (0.5 * wp_ver_size - yOffset) * tand(wp_PBA);

        % 計算等效眼睛
        E0 = [eyePosition(1:2),WDR*cosd(VVA)];
        E1 = E0 - [xOffset,yOffset,we_height];
        E2 = (rotx(-wp_PBA)^(-1) * E1')'; % 左手坐標系 wp_PBA --> -wp_PBA
        
        % 換算 VA
        [wdr_wp,theta_wp,phi_wp] = P2A(-E2(2),E2(1),E2(3));
        WDRLT = wdr_wp; VVALT = theta_wp; HVALT = phi_wp;
        HVALT = mod(360-HVALT,360); 

        % dummy plane 參數 assume "wp_PRA = 0"
        zOffset = moduleTop + 0.001 + we_height;
        dummyAlpha = wp_PBA;
        dummyBeta = 0;
        dummyGamma = 0;
        break
    end
    
    %% 接收器建立 %% For 20點分析: 每個Loop WDR不一樣 都要更新
    while (1)
    cprintf('key',"建立接收器......")
    % Dummy Plane: "D"; Receiver: "R"
    ReceiverDelete(lt); % 刪除當前接收器
    lt.Cmd("\V3D"); % 切到世界座標介面
    % dummy plane 建立
    lt.Cmd('DummyPlane ');
    lt.Cmd('XYZ'); % 虛擬面原點
    lt.Cmd(strcat(num2str(xOffset),',',num2str(yOffset),',',num2str(moduleTop+0.001)));
    lt.Cmd('XYZ'); % 虛擬面法線點 (與原點相減為法線)
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
    % 接收器 建立
    lt.Cmd(strcat('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"'));
    lt.Cmd('"Add Receiver"=');
    lt.Cmd('\Q');
    lt.Cmd('\O"LENS_MANAGER[1].ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last]"');    %Follow 2*N
    lt.Cmd('Name=R');
    lt.Cmd('Responsivity=Photometric '); % 單位 光通量
    %lt.Cmd('Responsivity=Radiometric '); % 單位 輻射通量
%     lt.Cmd('"Photometry Type"="Photometry Type B" '); % 方向:測光類型
    lt.Cmd('\Q');
    % 關閉向前模擬
    lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward Simulation]"');
    lt.Cmd('Enabled=No ');
    lt.Cmd('"Has Illuminance"=No ');
    lt.Cmd('"Has Intensity"=No ');
    lt.Cmd('\Q');
    % 向後模擬 + 空間輝度 設定
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
    elseif projectRec == 1 % 眼睛法向量與接收面垂直
        lt.Cmd(strcat('Long=',"0",' '));     % update each loop
        lt.Cmd(strcat('Lat=',"0",' '));      % update each loop
    end
    lt.Cmd('\Q');
    %% 20231019 seed 設定
    while (seed_firstTime == 1)
        lt.Cmd("\V3D"); % 切到世界座標介面
        lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].SIMULATIONS[ForwardAll]" ');
        lt.Cmd(strcat("StartingSeed=",num2str(seed)));
        lt.Cmd('\Q');
        seed_firstTime = 0;
    break
    end
    cprintf('err',"完成\n")
    break
    end
    disp(strcat("當前處理:"," WDR",num2str(WDR)," VVA",num2str(VVA_Ori)," HVA",num2str(HVA_Ori)," STA",num2str(systemTiltAngle)))
    if systemTiltAngle ~= 0 
        disp(strcat(" (等效角度: VVA",num2str(VVA)," HVA",num2str(HVA),")"))
    end

    %% 開始模擬
    while (1)
    cprintf('key',"LT 模擬中......")
    tStart = tic;
    lt.Cmd("\V3D");
    lt.Cmd("BeginAllSimulations");
    tEnd = toc(tStart);
    cprintf('err',"完成\n")
    break
    end
    %% 檢查 嘗試次數
    errActual = lt.DbGet("SURFACE_RECEIVER[R].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance].SPATIAL_LUMINANCE_MESH[Spatial Luminance Mesh]","ErrorAtPeak_Percent");
    cprintf('text',"峰值誤差: %.2f %% \n",round(errActual*100)/100)
    ARM = lt.DbGet("SURFACE_RECEIVER[R].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance]","Actual_Ray_Multiplier");
    if MRM < ARM;warning("最大光線嘗試次數 小於 實際光線嘗試次數");end
    %% 存取LT 檔案
    while (1)
    if saveLTFile==1
        cprintf('key',"[info] LT 存檔中......")
        backupFileName = strcat("LTFile Fig",num2str(N)," ",dateString," (Backup)");
        lt.SetOption('ShowFileDialogBox', 0);     % 自動存檔 不叫出存檔視窗 (必要)
        lt.Cmd('\VConsole');
        lt.Cmd(strcat("SaveAs """,fullfile(fullDirc,backupFileName),""""));
        lt.SetOption('ShowFileDialogBox', 1);     % 回復設定 (必要)
        cprintf('err',"完成\n")
        dips(strcat("[info] 檔名: ",backupFileName,".lts"));
    end
    break
    end
    %% 存取向後模擬Raw Datas & LTImage Photos
    while (1)
    if saveLTRaw==1
        cprintf('key',"LT Raw Data 存檔中......")
        meshkey = "SURFACE_RECEIVER[R].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance].SPATIAL_LUMINANCE_MESH[Spatial_Luminance_Mesh]";
        Xdim = lt.DbGet(meshkey, "X_Dimension");
        Ydim = lt.DbGet(meshkey, "Y_Dimension");
        DataArray=zeros(Ydim,Xdim);  %動態陣列
        [~,DataArray] = lt.GetMeshData(meshkey, DataArray(), "CellValue");
        DataArray=rot90(double(DataArray)); % 必要處理
        % 軸值
        xAxisFirst = str2double(string(lt.DbGet(meshkey,"XCellCenterAt",def,1)));
        xAxisFinal = str2double(string(lt.DbGet(meshkey,"XCellCenterAt",def,Xdim)));
        xAxisArray = linspace(xAxisFirst,xAxisFinal,Xdim);
        yAxisFirst = str2double(string(lt.DbGet(meshkey,"YCellCenterAt",def,1)));
        yAxisFinal = str2double(string(lt.DbGet(meshkey,"YCellCenterAt",def,Ydim)));
        yAxisArray = linspace(yAxisFirst,yAxisFinal,Ydim)';
        strAtA1 = {'Y\X'};
        % 寫Excel: 20220701
        % 檢查 改檔案是否已經存在: if yes --> delete then write
        % 檔案存在下 連續使用 'UseExcel',true 會出Bug 
        if isfile(excelFilepath);delete(excelFilepath);end
        writecell(strAtA1,excelFilepath,'Range','A1','AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Range','B1','AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(yAxisArray,excelFilepath,'Range','A2','AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(DataArray,excelFilepath,'Range','B2','AutoFitWidth',false); % 搭配 xlsx2png coding 格式
        cprintf('err',"完成\n")
        %% 轉圖檔
        cprintf('key',"匯出圖檔中......")
        maxValue = max(max(DataArray));
        minValue = min(min(DataArray));
        LTImage = uint8((DataArray./maxValue)*255);               % 0當最低值
        imwrite(LTImage,imageFilepath);
        cprintf('err',"完成\n")
    end
    break
    end
    cprintf('comment',"Loop完成\n")
    if lengthFiles ~= 1
        disp(strcat("花費時間: ",num2str(toc(tStartEachLoop))," 秒"))
    end
    % 嘗試 flush memory
    lt.Cmd("FlushUndoMemory");
    lt.Cmd("FlushDeletedEntityMemory");
    lt.Cmd("FlushSavedRayDataMemory");
    lt.Cmd("FlushAllMemory");
    pause(2)
end
cprintf('text',"程序完成\n")
disp(strcat("總花費時間: ",num2str(toc(tStartFromBegining))," 秒"))
%% Function
function [WDR,theta,phi] = Position2Angle(viewPointX,viewPointY,WDZ)
% Position2Angle: 直角坐標與球座標轉換 (By GY 20220613)
% < Function Handle Example >
% P2A = @Position2Angle;
% [vd,theta,phi] = P2A(X,Y,Z);
% X,Y: X Y coordinate; (右手坐標系即可)
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
function ReceiverDelete(lt) % 刪除當前所有DummyPlane (即包含接收器)
    lt.Cmd("\V3D");
    ObjectList = lt.DbList("COMPONENTS[1]", "PLANE_DUMMY_SURFACE");
    ObjectListSize = lt.ListSize(ObjectList); % 紀錄當前 Dummy Plane 總數
    for tt = 1 : ObjectListSize
        ObjectKey = lt.ListNext(ObjectList);
        ObjectName = lt.DbGet(ObjectKey, "NAME");
        lt.Cmd(strcat('\O',"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[",string(ObjectName),']'));
        lt.Cmd('Delete');
    end
    lt.ListDelete(ObjectList);
end
function sourceName = LightSourceSetup(lt) % 偵測當前所有光源
    lt.Cmd("\V3D");
    ObjectList = lt.DbList("SOURCES[Source_List]", "CUBE_SURFACE_SOURCE");
    ObjectListSize = lt.ListSize(ObjectList); % 紀錄當前 LightSource (Cube) 總數
    if ObjectListSize == 0
        error("錯誤: 當前 LT Model 沒有光源 (cube)")
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
        error("錯誤: 找不到可用的光源")
    end
    if countAvailable ~= 1
        error("錯誤: 除主光源, 其他光源請設定 1. 關閉光線可追跡性 2. 停用")
    end
    lt.ListDelete(ObjectList);
end
function checkExcelFileName(fullfilename)
    numtest = strlength(fullfilename);
    if numtest > 218
        error("存出 Excel 完整檔名字元數必須小於 218")
    end
end
% 當系統傾斜時 等效觀看角度
function [theta_polar_angle,phi_azimuthal_angle] = TiltAngle(WD,theta_polar_angle,phi_azimuthal_angle,SystemTiltAngle)
    %% 20230927 update
    % 紀錄 VVA 正負
    sign_VVA = abs(theta_polar_angle)/theta_polar_angle;
    
    % 得實際人眼位置
    WD_z=WD*cosd(theta_polar_angle); 
    ViewPoint_x=WD*sind(theta_polar_angle)*cosd(phi_azimuthal_angle);
    ViewPoint_y=WD*sind(theta_polar_angle)*sind(phi_azimuthal_angle);
    pointEye = [ViewPoint_x;ViewPoint_y;WD_z];
    pointEye_roty = roty(-SystemTiltAngle)*pointEye;
    
    % 反推 VVA (必為正值)
    theta_polar_angle = acosd(pointEye_roty(3)/WD);
    
    % 反推 HVA
    if pointEye_roty(1)~=0 % 非水平線
        if pointEye_roty(2) == 0 && pointEye_roty(1)< 0 % X 值 < 0, HVA 180
                phi_azimuthal_angle = 180;
        elseif pointEye_roty(2) == 0 && pointEye_roty(1)> 0 % X 值 > 0, HVA 0
                phi_azimuthal_angle = 0;
        else
            % atand: 鎖在 -90~90 
            phi_azimuthal_angle = atand(pointEye_roty(2)/pointEye_roty(1));
            if pointEye_roty(1) < 0 % atand 會算錯
               phi_azimuthal_angle = phi_azimuthal_angle + 180;
            end
        end
    else % 在水平線上，HVA等於原值 (正負由 VVA 正負決定) (EX: no prism DV HVA 方向)
        if sign_VVA < 0
            phi_azimuthal_angle = -phi_azimuthal_angle;
        elseif sign_VVA > 0
%             phi_azimuthal_angle;
        end
    end
    % HVA 控制在 +- 180 之間
    if phi_azimuthal_angle > 180
        phi_azimuthal_angle = phi_azimuthal_angle - 360;
    elseif phi_azimuthal_angle < -180
        phi_azimuthal_angle = phi_azimuthal_angle + 360;
    end
end