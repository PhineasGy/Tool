%% Original Author: BM %%
% 20220620: Modified by GY (project since 20220620)
% 用途: 13 點格柵感分析
% 確保 Dummy Plane 和 Receiver 的命名 (DX RX)
% 確保 有且只有一個 啟用+光線可追跡性開啟 的 CUBE_SURFACE_SOURCE
% Last Update: 20230926
% content: direct viewing VA bug fix
close all;clear;clc;
tStartFromBegining = tic;
%% 使用者輸入
LTID = 19460;
saveLTFile = 0;                                     % 存取 LT File開關
saveLTRaw = 1;                                      % 存取 LT excel, Fig.開關
chooseFolder = 1;                                   % 選擇root資料夾 (會在該資料夾新建LT還原資料夾)
    customLineFolder = '';                          % 資料夾名稱: "MRM5 LT還原 Custom..."
    customLineFile = '';                            % RawData/Png名稱: "(LT還原) II名 Custom..."
customDirectionGridApodizer = 1;                    % 光源角度分佈 (0:Lambertian，1:使用者自訂，-1:維持現狀不詢問)
seedPool = 1:5;                                     % random example: randi([1,100],[1,3]) 取三個 1-100 的 seed
NumReceiverArray = 1:13;                            % 1-13

% 13 點位置設定
% LT設定: (反向XY面) 向右x (HorSize)  向上y (VerSize)
moduleHorSize = 165.24;     % mm
moduleVerSize = 293.76;     % mm
moduleTop = 11.99;          % mm

% Receiver
buildRec = 1;                                       % 新建13點Receivers開關 (0:不重建)
    smoothFactor = 0;                               % 是否打開平滑. 0: 關閉, other: 3,5,7,...21 (限奇數)
    receiverSizeHor = 5; % mm
    receiverSizeVer = 5; % mm
    humanFactor = 1;
    verGridSize = 0.0765; % mm
    horGridSize = 0.0765; % mm
    horGridNum = "";
    verGridNum = "";
expectedERR = 0.02;
MRM = 20;                                           % 最大嘗試次數 (vs Actual Ray Multiplier ARM)
rayFactor = 1;                                      % 光線數額外加乘

% 眼睛參數 (向後空輝)
eyeMode = 0;                                        % -1 0 1 左中右眼
pupilSize = 2.5;                                    % (mm)
pupilBTDistance = 60;                               % IPD (mm)
% VA Array % 基於 中心眼 對 中心面板 % 可設定 PL_Array
viewingSetWDR = [400];
viewingSetVVA = [30];
viewingSetHVA = [0];                                
systemTiltAngle = 0;    

% Erosion Setup
erosionCS= 1;
structureElementRadius = 3;                         % 1 mm 代表幾個 Pixel (TBD)
structureElementHeight = 0;                         % offset
%% PreProcessing
while (1)
saveLTRawData = 1;     % 必須開啟
% 檢查 WDR VVA HVA 組數
if ~(length(viewingSetWDR)==length(viewingSetVVA)&&length(viewingSetWDR)==length(viewingSetHVA))
    error("VA組數必須相同");
end
dateString = datestr(now,'mm-dd-yyyy HH-MM');%% structure element
% erosion unit
se = offsetstrel('ball',structureElementRadius, structureElementHeight);
break
end
%% 跑數組結果 Loop
onlyCheckFirstTime = 0;
seedNum = length(seedPool);
setNum = length(viewingSetWDR);
CSMeanVAArray = cell(1,setNum);
for whichVANum = 1:setNum
tStartVASet = tic;
intensityArrayEachSeed = cell(1,seedNum); % 初始化
CSArrayEachSeed = cell(1,seedNum);
close all
WDR = viewingSetWDR(whichVANum);
VVA = viewingSetVVA(whichVANum);
HVA = viewingSetHVA(whichVANum);
% 考慮 tilt
VVA_Ori = VVA; % 紀錄初始角度用
HVA_Ori = HVA; % 紀錄初始角度用
[VVA,HVA] = TiltAngle(WDR,VVA,HVA,systemTiltAngle);
disp(strcat("處理中: ","WD",num2str(WDR)," VVA",num2str(VVA_Ori)," HVA",num2str(HVA_Ori)," STA",num2str(systemTiltAngle)))
%% seed Pool
for whichSeedNum = 1:seedNum
disp("------------")
tStartEachSeed = tic;
disp(strcat("Seed: ",num2str(seedPool(whichSeedNum))))
whichSeed = seedPool(whichSeedNum);  
onlyCheckFirstTime = onlyCheckFirstTime + 1;
%% 創建資料夾 (只在第一次建立)
while (1)
if onlyCheckFirstTime == 1
    if ~isempty(customLineFolder)
        customLineFolder = strcat(" ",customLineFolder);
    end
end
targetDirc = strcat("MRM",num2str(MRM)," 13點空輝 Seed ",num2str(whichSeed),customLineFolder);  %自定義資料夾名稱
rootFolder = string(rootFolder);
if isempty(rootFolder) || isequal(rootFolder,"")
    if onlyCheckFirstTime == 1 % 只選一次資料夾位子
        rootFolder = uigetdir("","選擇目標資料夾 (將在該資料夾中建立LT還原資料夾)");
        if rootFolder == 0;disp("系統停止");return;end
    end
elseif ~isstring(rootFolder)
    beep
    error("'rootFolder' variable should be either string or char type. (系統停止)")
end

lastwarn(''); % 重置 warning
fullDirc = fullfile(rootFolder,targetDirc);
mkdir(fullDirc);
msg = lastwarn;
if isequal('Directory already exists.',msg)
    disp("< 自定義資料夾已存在當前目錄 >")
end
break
end
%% 連結 LT (Citrix Version) % 只在第一次建立
while (onlyCheckFirstTime==1) 
def=System.Reflection.Missing.Value;
% ltcom64path=['C:\Program Files\Optical Research Associates\LightTools 9.1.1\Utilities.NET\LTCOM64.dll'];      %LTCOM64.dll路徑
ltcom64path=['C:\Program Files (x86)\Common Files\Optical Research Associates\LightTools\LTCOM64.dll'];      %LTCOM64.dll路徑_Critrix!!
asm=NET.addAssembly(ltcom64path);
lt=LTCOM64.LTAPIx;
lt.LTPID=LTID;
lt.UpdateLTPointer;
lt.Message('已連結到Matlab巨集');
break
end
%% seed
lt.Cmd("\V3D"); % 切到世界座標介面
lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].SIMULATIONS[ForwardAll]" ');
lt.Cmd(strcat("StartingSeed=",num2str(whichSeed)));
lt.Cmd('\Q');
%%
%% 讀取檔案: 角度分佈模式決定 (只處理一次)
while onlyCheckFirstTime == 1
% 面板光源名稱擷取
sourceName = LightSourceSetup(lt); % 取得當前光源名稱 (cube)
% 在光源"右面"切換切趾檔
source = strcat("CUBE_SURFACE_SOURCE[",sourceName,"]");
currentDG_first = string(lt.DbGet(strcat(source,".NATIVE_EMITTER[RightSurface]"),"Direction Apodizer Type"));
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
break
end
%% 13點計算
while (1)
% LT設定: 向右x (HorSize)  向上y (VerSize)
receiverPosition = zeros(25,3);
Wr = [1 5 7 8 9 12 13 14 17 18 19 21 25];
receiverZ = moduleTop + 0.001;
viewPointX = WDR*sind(VVA)*cosd(HVA); % 垂直方向VA
viewPointY = WDR*sind(VVA)*sind(HVA); % 水平方向VA
pupilCenterXY = [0;eyeMode*pupilBTDistance*0.5]; %以眼球中心為零點
% 兩眼旋轉
% 不旋轉條件: no Prism 掃HVA時 (HVA=90 && VVA ~=0)
if HVA ~= 90
    pupilCenterXYZTemp=[pupilCenterXY;0];
    pupilCenterXYZRotTemp = rotz(HVA) * pupilCenterXYZTemp;
    pupilCenterXY = pupilCenterXYZRotTemp(1:2); % 旋轉後向量
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
%% 人因設定 光線數設定
while (1)
if humanFactor == 1
    angRes = tand(1/120); % 人眼角分辨率
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
%% Receiver Build
while (1)
cprintf('key',"LT 模型建立中......")
ReceiverDelete(lt); % 刪除當前所有 Dummy Plane
sourceName = LightSourceSetup(lt); % 空間網格設為均勻
if buildRec==1
    for N = NumReceiverArray
        WDRLT = VALTList(Wr(N),1);
        VVALT = VALTList(Wr(N),2);
        HVALT = VALTList(Wr(N),3);
        lt.Cmd("\V3D"); % 切到世界座標介面
        % dummy plane 建立
        lt.Cmd('DummyPlane ');
        lt.Cmd('XYZ'); % 虛擬面原點
        lt.Cmd(strcat(num2str(receiverPosition(Wr(N),1)),',',num2str(receiverPosition(Wr(N),2)),',',num2str(receiverPosition(Wr(N),3))));
        lt.Cmd('XYZ'); % 虛擬面法線點 (與原點相減為法線)
        lt.Cmd(strcat(num2str(receiverPosition(Wr(N),1)),',',num2str(receiverPosition(Wr(N),2)),',',num2str(receiverPosition(Wr(N),3)+1)));
        lt.Cmd('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"');
        lt.Cmd(strcat('Name=D',num2str(Wr(N))));
        lt.Cmd('\Q');
        lt.DbSet(strcat("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D",num2str(Wr(N)),']'),"Width",receiverSizeHor);
        lt.DbSet(strcat("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D",num2str(Wr(N)),']'),"Height",receiverSizeVer);
        % 接收器 建立
        lt.Cmd(strcat('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"'));
        lt.Cmd('"Add Receiver"=');
        lt.Cmd('\Q');
        lt.Cmd('\O"LENS_MANAGER[1].ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last]"');
        lt.Cmd(strcat('Name=R',num2str(Wr(N))));
        lt.Cmd('Responsivity=Photometric ');
        lt.Cmd('\Q');
        % 關閉 向前模擬
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
    cprintf('err',"完成\n")
else % (不刪除接收器)
    ReceiverCheck(lt)
end    
break
end
%% Begin Simulation
while (1)
cprintf('key',"LT 模擬中......")
tStartSimulation = tic;
lt.Cmd("\V3D");
lt.Cmd("BeginAllSimulations");
tEndSimulation = toc(tStartSimulation);
cprintf('err',strcat("完成 (",num2str(tEndSimulation)," seconds)\n"))
break
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
%% 開啟向後模擬-空間輝度 (optional)
% lt.Cmd("\V3D");
% lt.Cmd(['LumViewSpatialLuminanceChart "R13 向後_模擬"']);
%% 存取資料
while (1)
if saveLTRawData==1
    cprintf('key',"寫入資料中......")
    dataArrayAll = cell(13,1);
    for N = NumReceiverArray
        meshKey = strcat("SURFACE_RECEIVER[R",num2str(Wr(N)),"].BACKWARD_SPATIAL_LUMINANCE[Spatial_Luminance].SPATIAL_LUMINANCE_MESH[Spatial_Luminance_Mesh]");
        Xdim = lt.DbGet(meshKey, "X_Dimension");
        Ydim = lt.DbGet(meshKey, "Y_Dimension");
        imageFilepath = fullfile(fullDirc,strcat("R",num2str(Wr(N)),".png"));   % 存結果 位置
        xlsstr = strcat("Result_VVA",num2str(VVA),"_HVA",num2str(HVA),"_WDR",num2str(WDR),"_MRM",num2str(MRM),"_",dateString,".xlsx");
        excelFilepath = fullfile(fullDirc,xlsstr);  % 存結果 位置
        %% 輸出Excel
        % 網格值
        dataArray=zeros(Ydim,Xdim);  %動態陣列
        [~,dataArray] = lt.GetMeshData(meshKey, dataArray(), "CellValue");
        dataArray=rot90(double(dataArray)); % 必要處理
        dataArrayAll{N} = dataArray;
        % 軸值
        xAxisFirst = str2double(string(lt.DbGet(meshKey,"XCellCenterAt",def,1)));
        xAxisFinal = str2double(string(lt.DbGet(meshKey,"XCellCenterAt",def,Xdim)));
        xAxisArray = linspace(xAxisFirst,xAxisFinal,Xdim);
        yAxisFirst = str2double(string(lt.DbGet(meshKey,"YCellCenterAt",def,1)));
        yAxisFinal = str2double(string(lt.DbGet(meshKey,"YCellCenterAt",def,Ydim)));
        yAxisArray = linspace(yAxisFirst,yAxisFinal,Ydim)';
        strAtA1 = {'Y\X'};
        writecell(strAtA1,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','A1','UseExcel',true,'AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','B1','UseExcel',true,'AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(yAxisArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','A2','UseExcel',true,'AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(dataArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','B2','UseExcel',true,'AutoFitWidth',false); % 搭配 LT 寫出格試
        %% 轉圖檔
        maxValue = max(max(dataArray));
        minValue = min(min(dataArray));
        LTImage = uint8((dataArray./maxValue)*255);               % 0當最低值
        imwrite(LTImage,imageFilepath);
    end
    cprintf('err',"完成\n")
end 
break
end
%% 記錄峰值誤差 (20220609 By GY)
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
        warning("有接收面沒有達至設定光線數")
    end
end

break
end
%% 格柵感分析
while (1)
cprintf('key',"格柵感分析......")
LTWDR = zeros(5,10);
intensityArray = zeros(2,13);
Wr = [1 5 7 8 9 12 13 14 17 18 19 21 25];
Datalist = zeros(5,13);
Datalist(1,:) = Wr;
F_each = cell(13,1);
maxIntensityArray = zeros(5,5);
for temptemp = 1:13 % 初始化
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
        imageFilepath = fullfile(fullDirc,strcat("R",num2str(Wr(k)),"_Erosion.png"));   % 存結果 位置
        imwrite(XX,imageFilepath);
        picTargetErodeArray{row,column} = XX;
    end

    %% 累積函數找10% & 90%
    H_forobserve=histogram(XX);
    H=histogram(XX,'Normalization','cdf');
    edge=[H.BinLimits(1):0.5:H.BinLimits(end)];
    H=histogram(XX,'Normalization','cdf','BinEdges',edge);
    CumData_Y=H.Values;
    CumData_X=zeros(1,length(H.Values));
    for jj=1:length(H.BinEdges)-1
        CumData_X(jj)=0.5*(H.BinEdges(jj)+H.BinEdges(jj+1));
    end
    x_const=0:5:255;  %sim 小的值可能需調整 0.1 大的可以5  !!!!!
    y_const=x_const;
    y_const(:)=0.1;
    Imin=polyxpoly(x_const,y_const,CumData_X,CumData_Y);
    y_const(:)=0.9;
    Imax=polyxpoly(x_const,y_const,CumData_X,CumData_Y);
    %%
    CR=(Imax-Imin)/(Imax+Imin);   %輸出對比度
    CSF1=1/CR;
    LTWDR(Wr(k)) = roundn(CSF1,-1);
    intensityArray(1,k)=Imax;
    intensityArray(2,k)=Imin;
    
    W_non0 = LTWDR(LTWDR(:,1:5)>0);
    W_non0_AVG=round(sum(W_non0)/13,1);

    %% Imax Ratio 20220609 By GY
    [row,column] = ind2sub([5,5],Wr(k));
    maxIntensityArray(row,column) = Datalist(2,k);
    if k == NumReceiverArray(1) % 紀錄第一組中心點強度 (位置13) (第7點)
        maxIntensitySpecific = maxIntensityArray(row,column);
    end
end
if erosionCS == 1
    disp("CS Value (after erosion)");
elseif erosionCS == 0
    disp("CS Value");
end
disp(LTWDR);
disp(strcat("平均: ",num2str(W_non0_AVG)));
maxIntensityArrayNormal2Each = maxIntensityArray/max(maxIntensityArray(:));
maxIntensityArrayNormal2Specific = maxIntensityArray/maxIntensitySpecific;

%% Summary
cprintf('key',"寫入Summary中......")
F_zero=zeros(Picc(1),Picc(2));
Total_F=[F_each{1},F_zero,F_zero,F_zero,F_each{12};
        F_zero,F_each{3},F_each{6},F_each{9},F_zero;
        F_zero,F_each{4},F_each{7},F_each{10},F_zero;
        F_zero,F_each{5},F_each{8},F_each{11},F_zero;
        F_each{2},F_zero,F_zero,F_zero,F_each{13}];

figure;  imshow(Total_F,[0 255]);
imwrite(Total_F,fullfile(fullDirc,"13點圖.png"));

writematrix(LTWDR,excelFilepath,'Sheet','Summary!!','UseExcel',true,'AutoFitWidth',false);
writematrix(W_non0_AVG,excelFilepath,'Sheet','Summary!!','Range','A7','UseExcel',true,'AutoFitWidth',false);
txt={'Max';'Min';'Min/Max(%)';'CS (100%-0%)'};
writecell(txt,excelFilepath,'Sheet','Summary!!','Range','A12','UseExcel',true,'AutoFitWidth',false);
writematrix(Datalist,excelFilepath,'Sheet','Summary!!','Range','B11','UseExcel',true,'AutoFitWidth',false);
writecell({'峰值誤差(%)'},excelFilepath,'Sheet','Summary!!','Range','P3','UseExcel',true,'AutoFitWidth',false);
writecell({'實際接收光線數'},excelFilepath,'Sheet','Summary!!','Range','W3','UseExcel',true,'AutoFitWidth',false);
writecell({'預期接收光線數'},excelFilepath,'Sheet','Summary!!','Range','W2','UseExcel',true,'AutoFitWidth',false);
writecell({'最大嘗試次數'},excelFilepath,'Sheet','Summary!!','Range','AD3','UseExcel',true,'AutoFitWidth',false);
writecell({'最大光強 normal2self'},excelFilepath,'Sheet','Summary!!','Range','P10','UseExcel',true,'AutoFitWidth',false);
writecell({'最大光強 normal2all'},excelFilepath,'Sheet','Summary!!','Range','W10','UseExcel',true,'AutoFitWidth',false);
writematrix(errorPeakPercentageArray,excelFilepath,'Sheet','Summary!!','Range','Q3','UseExcel',true,'AutoFitWidth',false);
writematrix(actualRayHitArray,excelFilepath,'Sheet','Summary!!','Range','X3','UseExcel',true,'AutoFitWidth',false);
writematrix(rayNum,excelFilepath,'Sheet','Summary!!','Range','X2','UseExcel',true,'AutoFitWidth',false);
writematrix(actualRayMultiplierArray,excelFilepath,'Sheet','Summary!!','Range','AE3','UseExcel',true,'AutoFitWidth',false);
writematrix(maxIntensityArrayNormal2Each,excelFilepath,'Sheet','Summary!!','Range','Q10','UseExcel',true,'AutoFitWidth',false);
writematrix(maxIntensityArrayNormal2Specific,excelFilepath,'Sheet','Summary!!','Range','X10','UseExcel',true,'AutoFitWidth',false);
cprintf('err',"完成\n")
break
end
%% error peak 相關
while (1)
if errorCheck == 1
    disp(strcat("max(最大嘗試次數) = ",num2str(max(actualRayMultiplierArray(:)))));
else
    disp("所有接收面達至設定光線數.")
end
disp("峰值誤差")
disp(errorPeakPercentageArray)
break
end
%% 紀錄 Imax Imin (13點) (For CS<mean>)
intensityArrayEachSeed{whichSeedNum} = intensityArray;
CSArrayEachSeed{whichSeedNum} = LTWDR(:,1:5);

if erosionCS == 1
picTargetErodeCombine = cell2mat(picTargetErodeArray);
imageFilepath = fullfile(fullDirc,"combination_Erosion.png");   % 存結果 位置
imwrite(picTargetErodeCombine,imageFilepath);
end
tEndEachSeed = toc(tStartEachSeed);
disp(strcat("process complete (each seed): ",num2str(tEndEachSeed)," seconds."))
end % seed 結束
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
end % VA 組 結束
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
function ReceiverCheck(lt) % 檢查當前所有 Dummy Plane 和 接收器 命名是否正確
    % 檢查 Dummy Plane
    DummyPlanePool = ["D1","D5","D7","D8","D9","D12","D13","D14","D17","D18","D19","D21","D25"];
    ObjectList = lt.DbList("COMPONENTS[1]", "PLANE_DUMMY_SURFACE");
    ObjectListSize = lt.ListSize(ObjectList); % 紀錄當前 Dummy Plane 總數
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
    disp("Dummy Plane 確認完畢")
    lt.ListDelete(ObjectList);
    % 檢查 Receiver
    ReceiverPool = ["R1","R5","R7","R8","R9","R12","R13","R14","R17","R18","R19","R21","R25"];
    ObjectList = lt.DbList("RECEIVERS[Receiver_List]", "SURFACE_RECEIVER");
    ObjectListSize = lt.ListSize(ObjectList); % 紀錄當前 Dummy Plane 總數
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
    disp("Receiver 確認完畢")
    lt.ListDelete(ObjectList);
end
function sourceName = LightSourceSetup(lt) % 光源設定
    % 目前功能:
    % 1. 檢查有效光源 (cube)
    % 2. 空間網格改為均勻
    %%%%%%
    % 1. 檢查光源
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
    
    % 2. 空間網格設定為均勻 (13點Only)
    source = strcat("CUBE_SURFACE_SOURCE[",sourceName,"]");
    lt.Cmd("\V3D");
    lt.Cmd(strcat("\O",source,".NATIVE_EMITTER[RightSurface]"));
    lt.Cmd('"Surface Apodizer Type"=Uniform ');
    lt.Cmd('\Q');

    % 最後清掉 object
    lt.ListDelete(ObjectList);
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