%% Original Author: BM %%
% 20220708: Modified by GY (project since 20220615)
% 用途: 極限解析20點
% 用途: LT RP回追
% Last Update: 20230926
% 20230926: direct viewing VA bug fix
close all;clear;clc;
tStartFromBegining = tic;
%% 使用者輸入
LTID = 31408;
saveLTFile = 0;                                     % 存取 LT File開關
saveLTRaw = 1;                                      % 存取 LT excel, Fig.開關
% root folder: will create folder here
rootFolder = "";                                    % 母資料夾決定. "": choose folder; cd: current folder (不用引號); "abc/def...": other path
    customLineFolder = '';                          % 結果資料夾額外名稱: "MRM5 LT還原 Custom..."
    customLineFile = '';                            % 結果檔案額外名稱: "(LT還原) II名 Custom..."
GLCritical = 5;                                     % 影像二值化處理 (大於該值以上 == 255)

% 接收器參數 (每次刪掉重建)
% 注意網格總數不得超過 4200000 (LT限制)
smoothFactor = 0;                                   % 使用開啟平滑. 0: 關閉, other: 3,5,7,...21 (限奇數)
xOffset = 0;                                        % Receiver LT_X offset (Hor) (往右為正)
yOffset = 0;                                        % Receiver LT_Y offset (Ver) (往上為正)
zOffset = 0.1;                                      % 面板 " 接收面 "位置 (mm)
receiverSizeHor = 165.24;                           % mm (if split on: 總接收面大小)
receiverSizeVer = 293.76;                           % mm (if split on: 總接收面大小)
horGridSize = "";                                   % mm % set to "" if dont needed (if split on: 總接收面大小)
verGridSize = "";                                   % mm % set to "" if dont needed (if split on: 總接收面大小)
horGridNum = 2160;                                  % set to "" if dont needed (if split on: 總接收面大小)
verGridNum = 3840;                                  % set to "" if dont needed (if split on: 總接收面大小)
expectedERR = 0.1;                                  % 預期峰值誤差 % (向前模擬僅供參考)
rayFactor = 1;                                      % 光線數額外加乘

% 結果影像是否調整大小
isResultResized = 0;
    resultResizeHorNum = 2160;                      % 調整後影像水平向大小
    resultResizeVerNum = 3840;                      % 調整後影像垂直向大小

% 是否以 receiverSize 作分割 (保持網格大小)
isReceiverSplit = 1;                                % 請確保拼接時沒有小數點情形導至誤差 EX: 3840/3 拼接有誤差
    receiverSplitNumHor = 2;                        % 水平向分割數量
    receiverSplitNumVer = 2;                        % 垂直向分割數量

% 眼睛參數 (決定光源位置)
buildLightSouce = 1;                                % 是否重建光源
moduleTop = 7.62;                                   % 系統最頂 Z 座標 (mm) 
eyeMode = 0;                                        % -1 0 1 左中右眼
pupilSize = 15;                                     % mm
pupilBTDistance = 60;                               % IPD mm
aimAreaHor = 165.24;                                % 光源定位範圍大小 (水平)
aimAreaVer = 293.76;                                % 光源定位範圍大小 (垂直)
aimCenterHor = 0;                                   % 光源定位範圍中心 (水平)
aimCenterVer = 0;                                   % 光源定位範圍中心 (垂直)
aimCenterZ = moduleTop;

% 中心單眼 的 VA 資訊
WDR = 500;                                          % Based on 0 Offset % (中心眼對中心面板)
VVA = 35;                                           % Based on 0 Offset % (中心眼對中心面板)
HVA = 0;                                            % Based on 0 Offset % (中心眼對中心面板)
systemTiltAngle = 5;                                % 將 VA 轉換為等效角度

% display info for image
dateStringOn = 1;
displayInfo = 1; % VA, STA, PS, eye, IPD
%% 光線數預估 其他參數確認
while (1)
if dateStringOn == 1
    dateString = strcat("_",datestr(now,'mm-dd-yyyy HH-MM'));
elseif dateStringOn == 0
    dateString = "";
end
switch eyeMode;case -1;eyeString = "leftEye";case 0;eyeString = "monoEye";case 1;eyeString = "rightEye";end
imageString = strcat("_WDR",num2str(WDR),"_VVA",num2str(VVA),"_HVA",num2str(HVA),...
    "_STA",num2str(systemTiltAngle),"_PS",num2str(pupilSize),"_",eyeString,"_IPD",num2str(pupilBTDistance));
disp("條件:")
disp(imageString)
if displayInfo == 0;imageString = "";end

smoothFactorPool = [0,3,5,7,9,11,13,15,17,19,21];
if ~any(smoothFactor==smoothFactorPool)
    error("平滑參數設定有誤")
end

if ~isequal(customLineFile,"")
    customLineFile = strcat("_",customLineFile);
end

% 考慮 tilt
VVA_Ori = VVA; % 紀錄初始角度用
HVA_Ori = HVA; % 紀錄初始角度用
[VVA,HVA] = TiltAngle(WDR,VVA,HVA,systemTiltAngle);

% 光線數 網格數 檢查
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
        error('請保持 網格數 or 網格大小 擇一為""');
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
        error('請保持 網格數 or 網格大小 擇一為""');
    end
end

if actualHGN * actualVGN >= 4200000
    error(strcat("錯誤: 網格中的單格數 (",num2str(horGridNum)," x ",num2str(verGridNum),") 超過  4200000 目前的限制"))
end

ERR = expectedERR;
rayNum = (actualHGN * actualVGN)/ERR^2;
rayNum = rayNum * rayFactor;
eyePosLT = EyePosition(WDR,VVA,HVA,eyeMode,pupilBTDistance,moduleTop); % 得光源位置

break
end
%% 創建資料夾
while (1)
if ~isempty(customLineFolder)
    customLineFolder = strcat("_",customLineFolder);
end
targetDirc = strcat("LTRP回追 ",customLineFolder);  %自定義資料夾名稱
rootFolder = string(rootFolder);
if isempty(rootFolder) || isequal(rootFolder,"")
    rootFolder = uigetdir("","選擇目標資料夾 (將在該資料夾中建立 LT 回追資料夾)");
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
%% 演算開始
cprintf("-------------------------------\n")
ReceiverDelete(lt); % 刪除當前所有接收器
%% 光源建立 %%
if buildLightSouce == 1
    checkBLS = 1;
elseif buildLightSouce == 0
    checkBLS = 0;
end
while checkBLS
SourceDelete(lt); % 刪除當前所有Disk光源 + 關閉Cube光源
cprintf('key',"建立光源......") 
lt.Cmd("\V3D"); % 切到世界座標介面
% Disk Source 建立
lt.Cmd('DiskSource ');
lt.Cmd('XYZ'); % 光源中心位置
lt.Cmd(strcat(num2str(eyePosLT(1)),',',num2str(eyePosLT(2)),',',num2str(eyePosLT(3))));
lt.Cmd('XYZ'); % 半徑位置
lt.Cmd(strcat(num2str(eyePosLT(1)+pupilSize),',',num2str(eyePosLT(2)),',',num2str(eyePosLT(3))));
lt.Cmd('XYZ'); % 法向量
lt.Cmd(strcat(num2str(eyePosLT(1)),',',num2str(eyePosLT(2)),',',num2str(eyePosLT(3)-1)));
lt.Cmd('\O"DISK_SOURCE[@Last]"');
lt.Cmd('Name=E');
lt.Cmd('"Power Extent"="Aim region" '); % 測量於定位範圍
lt.Cmd('"Aim Entity Type"="Aim Area" '); % 定位區域
lt.Cmd('"Power Units"=Photometric '); % 光通量
lt.Cmd('\Q');
% 找到其定位區域位址
ObjectList = lt.DbList("DISK_SOURCE[@Last]", "CIRCULAR_AIM_AREA");
ObjectListSize = lt.ListSize(ObjectList); % 紀錄當前 Source 總數
ObjectKey = lt.ListNext(ObjectList);
ObjectName = lt.DbGet(ObjectKey, "NAME");
lt.ListDelete(ObjectList);
lt.Cmd(strcat('\O"DISK_SOURCE[@Last].CIRCULAR_AIM_AREA[',string(ObjectName),']"'));
lt.Cmd('"Element Shape"=Rectangular '); % 定位區域形狀
lt.Cmd('\Q');
lt.Cmd(strcat('\O"DISK_SOURCE[@Last].RECTANGULAR_AIM_AREA[',string(ObjectName),']"'));
lt.Cmd(strcat('X=',num2str(aimCenterHor))); % 定位區域座標大小
lt.Cmd(strcat('Y=',num2str(aimCenterVer))); % 定位區域座標大小
lt.Cmd(strcat('Z=',num2str(aimCenterZ))); % 定位區域座標大小
lt.Cmd(strcat('Width=',num2str(aimAreaHor))); % 定位區域座標大小
lt.Cmd(strcat('Height=',num2str(aimAreaVer))); % 定位區域座標大小
lt.Cmd('\Q');
cprintf('err',"完成\n")
break
end
%% 建立接收器位置矩陣 (尚未考慮Offset)
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
            receiverPositionXYArray{ii,jj} = [PositionEach(2);-PositionEach(1)]; % ML座標 > LT座標
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
%% 接收器迴圈
imageArray = cell(receiverSplitNumVer,receiverSplitNumHor);
for whichReceiver = 1 : totalNumReceiver 
    tStartEachLoop = tic;
    cprintf("-------------------------------\n")
    disp(strcat("當前處理:"," WDR",num2str(WDR)," VVA",num2str(VVA_Ori)," HVA",num2str(HVA_Ori)," STA",num2str(systemTiltAngle)," 接收器 ",num2str(whichReceiver)))
    [row,column] = ind2sub([receiverSplitNumVer,receiverSplitNumHor],whichReceiver);
    receiverPositionXYIncludingOffset = receiverPositionXYArray{row,column} + [xOffset;yOffset];
    rPXYIO = receiverPositionXYIncludingOffset;
    while (1)
    cprintf('key',"建立接收器......")
    % Dummy Plane: "D"; Receiver: "R"
%     ReceiverDisable(lt) % 先停用之前的接收器 (發現並不會記錄原有資料)
    ReceiverDelete(lt); % 刪除當前所有接收器
    lt.Cmd("\V3D"); % 切到世界座標介面
    % dummy plane 建立
    lt.Cmd('DummyPlane ');
    lt.Cmd('XYZ'); % 虛擬面原點
    lt.Cmd(strcat(num2str(rPXYIO(1)),',',num2str(rPXYIO(2)),',',num2str(zOffset-0.001)));
    lt.Cmd('XYZ'); % 虛擬面法線點 (與原點相減為法線)
    lt.Cmd(strcat(num2str(rPXYIO(1)),',',num2str(rPXYIO(2)),',',num2str(zOffset)));
    lt.Cmd('\O"PLANE_DUMMY_SURFACE[@Last]"');
    lt.Cmd(strcat('Name=D',num2str(whichReceiver)));
    lt.Cmd('\Q');
    lt.DbSet("PLANE_DUMMY_SURFACE[@Last]","Width",receiverSizeUnitHor*2);
    lt.DbSet("PLANE_DUMMY_SURFACE[@Last]","Height",receiverSizeUnitVer*2);
    % 接收器 建立
    lt.Cmd(strcat('\O"PLANE_DUMMY_SURFACE[@Last]"'));
    lt.Cmd('"Add Receiver"=');
    lt.Cmd('\Q');
    lt.Cmd('\O"SURFACE_RECEIVER[@Last]"');
    lt.Cmd(strcat('Name=R',num2str(whichReceiver)));
    lt.Cmd('Responsivity=Photometric '); % 單位 光通量
    lt.Cmd('\Q');
    lt.Cmd('\O"SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward Simulation]"'); % 向前模擬
    lt.Cmd('"Has Intensity"=No '); % 關閉強度分析
    lt.Cmd('"Save Ray Data"=No '); % 關閉儲存光線
    lt.Cmd('\Q');
    lt.Cmd('\O"SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward Simulation].ILLUMINANCE_MESH[Illuminance Mesh]"');
    lt.Cmd(strcat('"X Dimension"=',num2str(actualHGN))); % 設定網格大小X (Hor)
    lt.Cmd(strcat('"Y Dimension"=',num2str(actualVGN))); % 設定網格大小Y (Ver)
    if smoothFactor == 0
        lt.Cmd('"Do Noise Reduction"=No '); % 取消平滑
    else
        lt.Cmd(strcat('"Kernel Size N"=',num2str(smoothFactor))); % 設定平滑常數
    end
    lt.Cmd('\Q');
    lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].SIMULATIONS[ForwardAll]" '); % 向前模擬總覽
    lt.Cmd('Enabled=Yes '); % 啟用向前模擬
    lt.Cmd(strcat('MaxProgress=',num2str(rayNum))); % 設定光線數 (向前:不符合誤差公式?)
    lt.Cmd('\Q');
    cprintf('err',"完成\n")
    break
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
    errActual = lt.DbGet("SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward_Simulation].ILLUMINANCE_MESH[Illuminance_Mesh]","ErrorAtPeak_Percent");
    cprintf('text',"峰值誤差: %.2f %% \n",round(errActual*100)/100)
    %% 存取向後模擬Raw Datas & LTImage Photos
    while (1)
    if saveLTRaw==1
        cprintf('key',"LT Raw Data 存檔中......")
        meshkey = "SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward_Simulation].ILLUMINANCE_MESH[Illuminance_Mesh]";
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
        excelFileName = strcat("(LTRP) R",num2str(whichReceiver),customLineFile,".xlsx");
        excelFilepath = fullfile(fullDirc,excelFileName);
        if isfile(excelFilepath);delete(excelFilepath);end
        writecell(strAtA1,excelFilepath,'Range','A1','AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Range','B1','AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(yAxisArray,excelFilepath,'Range','A2','AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(DataArray,excelFilepath,'Range','B2','AutoFitWidth',false); % 搭配 xlsx2png coding 格式
        cprintf('err',"完成\n")
        %% 轉圖檔
        imageFileName = strcat("(LTRP) (Raw) R",num2str(whichReceiver),customLineFile,".png");
        imageFilepath = fullfile(fullDirc,imageFileName);
        cprintf('key',"匯出圖檔中......")
        maxValue = max(max(DataArray));
        minValue = min(min(DataArray));
        LTImage = uint8((DataArray./maxValue)*255);               % 0當最低值
        if isResultResized == 1
            LTImage = imresize(LTImage,[resultResizeVerNum,resultResizeHorNum]);
        end
        [imageSizeRow,imageSizeColumn] = size(LTImage); % comibine 用途
        imwrite(LTImage,imageFilepath);
        imageArray{row,column} = LTImage;
        cprintf('err',"完成\n")
    end
    break
    end
    cprintf('comment',"Loop %d 完成\n",whichReceiver)
    if totalNumReceiver ~= 1
        disp(strcat("花費時間: ",num2str(toc(tStartEachLoop))," 秒"))
    end
end
%% 存取LT 檔案
while (1)
if saveLTFile==1
    cprintf('key',"LT 存檔中......")
    backupFileName = strcat("LTFile ",dateString," (Backup)");
    lt.SetOption('ShowFileDialogBox', 0);     % 自動存檔 不叫出存檔視窗 (必要)
    lt.Cmd('\VConsole');
    lt.Cmd(strcat("SaveAs """,fullfile(fullDirc,backupFileName),""""));
    lt.SetOption('ShowFileDialogBox', 1);     % 回復設定 (必要)
    cprintf('err',"完成\n")
end
break
end
% 合成全影像 % 以GLCritical作二值化
while (1)
if isReceiverSplit == 1 && saveLTRaw == 1
    % Raw 影像 (結合)
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
    % 二值影像 (結合)
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
cprintf('text',"程序完成\n")
disp(strcat("總花費時間: ",num2str(toc(tStartFromBegining))," 秒"))
%% Function
function ReceiverDelete(lt) % 刪除當前所有DummyPlane (即包含接收器)
    lt.Cmd("\V3D");
    ObjectList = lt.DbList("COMPONENTS[1]", "PLANE_DUMMY_SURFACE");
    ObjectListSize = lt.ListSize(ObjectList); % 紀錄當前 Dummy Plane 總數
    for tt = 1 : ObjectListSize
        ObjectKey = lt.ListNext(ObjectList);
        ObjectName = lt.DbGet(ObjectKey, "NAME");
        lt.Cmd(strcat('\O',"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[",string(ObjectName),']'));
        lt.Cmd('Delete');
        lt.Cmd('\Q');
    end
    lt.ListDelete(ObjectList);
end
% function ReceiverDisable(lt) % 將先前接收器停用 (保留每片接收器)
%     lt.Cmd("\V3D");
%     ObjectList = lt.DbList("RECEIVERS[Receiver List]", "SURFACE_RECEIVER");
%     ObjectListSize = lt.ListSize(ObjectList); % 紀錄當前 Dummy Plane 總數
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
function SourceDelete(lt) % 刪除當前所有Disk Source (關閉 Cube 光源) (目前尚無法同時找到所有種類光源)
    lt.Cmd("\V3D");
    % 刪除 DiskSource
    ObjectList = lt.DbList("SOURCES[Source_List]", "DISK_SOURCE");
    ObjectListSize = lt.ListSize(ObjectList); % 紀錄當前 Source 總數
    for tt = 1 : ObjectListSize
        ObjectKey = lt.ListNext(ObjectList);
        ObjectName = lt.DbGet(ObjectKey, "NAME");
        lt.Cmd(strcat('\ODISK_SOURCE[',string(ObjectName),']'));
        lt.Cmd('Delete');
        lt.Cmd('\Q');
    end
    lt.ListDelete(ObjectList);
    % 關閉 Cube Source (關閉光線可追跡性 停用)
    ObjectList = lt.DbList("SOURCES[Source_List]", "CUBE_SURFACE_SOURCE");
    ObjectListSize = lt.ListSize(ObjectList); % 紀錄當前 Source 總數
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
    
    % 中心單眼位置 (WDZ: z From systemTop)
    WDZ = WDR*cosd(VVA); 
    viewPointX = WDR*sind(VVA)*cosd(HVA);
    viewPointY = WDR*sind(VVA)*sind(HVA);

    % 兩眼中心旋轉
    eyeVector = [0;eyeMode*pupilBTDistance*0.5]; % 左右眼水平張開 沒旋轉

    % 不旋轉條件: no Prism 掃 HVA 時 (HVA=90 && VVA ~=0)
    if HVA~=90
        eyeVector3temp = [eyeVector;0];
        eyeVector3 = rotz(HVA) * eyeVector3temp;
        eyeVector = eyeVector3(1:2);
    end
    eyePos = [viewPointX; viewPointY; moduleTop + WDZ] + [eyeVector;0]; % ML坐標系
    
    % 令 Receiver Z = 0 + WD_z; 面板中心為 (x =0 y=0)
    eyePosLT = [eyePos(2);-eyePos(1);eyePos(3)];
    
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