%% 13 點角輝 巨集 %%
% Author: GY (project since 20220614) (Ref by BM)
% 確保 Dummy Plane 和 Receiver 的命名 (DX RX)
% 不做光源設定 (請先自行設定完成)
% Last Update: 20240320
% content: first git version, 
close all;clear;clc;
tic
%% 參數設定
% options
LTID = 13896;           % 同時開多個LT，指定用
saveLTFile = 0;         % 存取 LT File開關
saveLTRawData = 1;      % 存取 LT excel, Fig.開關
buildRec = 1;           % 新建 13 點 Receivers 開關 (0:不須重建)
chooseFolder = 1;       % 資料夾存放處
customLineFolder = '';    % 目標資料夾額外命名
rayFactor = 1;          % 光線數額外加乘

% parameters setup
% LT設定: (反向XY面) 向右x (HorSize)  向上y (VerSize)
horSize = 165.24;   % For 13 Point Position
verSize = 293.76;   % For 13 Point Position
moduleTop = 0.33;   % mm (接收面會自行 = 0.001mm)
% Receiver (type C)
receiverSizeHor = 5; % mm (Also affect 14th Receiver)
receiverSizeVer = 5; % mm (Also affect 14th Receiver)
ALMHalfSize = 2.5; % mm 計量器大小半徑 (建議 < 接收器大小/2)

LgridNum = 360;         % (Also affect 14th Receiver)
VgridNum = 90;          % (Also affect 14th Receiver)
expectedERR = 0.08; % 預期峰值誤差

MRM = 20;           % 最大嘗試次數 (vs Actual Ray Multiplier ARM)

% Debug 用 %
NumReceiverArray = 1:13; % 1-13, 14: 自訂(預設為中心) % 14: 單片(自訂位置) D38 R38  
    xOffset = 0;         % work when NumReceiverArray = 14
    yOffset = 0;         % work when NumReceiverArray = 14

% excel write 處理 %
dateString = datestr(now,'mm-dd-yyyy HH-MM');
xlsstr = strcat("Result_13點向後角輝_MRM ",num2str(MRM),"_",dateString,".xlsx");

%% 參數前處理
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

% 計量器 Check
if ALMHalfSize > min([receiverSizeHor,receiverSizeVer]*0.5)
    beep
    warnAns = questdlg('(Warning) 計量器直徑 大於 接收器大小','Warning','continue','cancel','cancel');
    switch warnAns
        case 'continue'
        otherwise
            disp("System Stopped.")
            return
    end
end
break
end
%% 創建資料夾
while (1)
if ~isempty(customLineFolder)
    customLineFolder = strcat("_",customLineFolder);
end
targetDirc = strcat("MRM",num2str(MRM)," 13點角輝",customLineFolder);  %自定義資料夾名稱
if chooseFolder == 0 % 在Matlab當前資料夾建立 13點角輝資料夾
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
    disp("< 自定義資料夾已存在當前目錄 >")
end
excelFilepath = fullfile(fullDirc,xlsstr);
checkExcelFileName(excelFilepath);
break
end
%% 連結 LT (Citrix Version)
while (1)
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
%% 13點計算 (+1:第14點)
while (1)
% LT設定: 向右x (HorSize)  向上y (VerSize)
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
%% 峰值誤差與光線數估計
while (1)
ERR = expectedERR;
rayNum = (LgridNum * VgridNum)/ERR^2;
rayNum = rayNum * rayFactor;
break
end
%% Receiver Build
while (1)
if buildRec==1
    cprintf('key',"LT 模型建立中......")
    ReceiverDelete(lt);     % 刪除當前所有 Dummy Plane
%     sourceName = LightSourceSetup(lt);   % 空間網格設為均勻
    for N = NumReceiverArray
        lt.Cmd("\V3D"); % 切到世界座標介面
        % dummy plane 建立
        lt.Cmd('DummyPlane ');
        lt.Cmd('XYZ'); % 虛擬面原點
        lt.Cmd(strcat(num2str(listPostion(Wr(N),1)),',',num2str(listPostion(Wr(N),2)),',',num2str(listPostion(Wr(N),3))));
        lt.Cmd('XYZ'); % 虛擬面法線點 (與原點相減為法線)
        lt.Cmd(strcat(num2str(listPostion(Wr(N),1)),',',num2str(listPostion(Wr(N),2)),',',num2str(listPostion(Wr(N),3)+1)));
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
        lt.Cmd('Responsivity=Photometric '); % 單位 光通量
        %lt.Cmd('Responsivity=Radiometric '); % 單位 輻射通量
        lt.Cmd('"Photometry Type"="Photometry Type C" '); % 方向:測光類型
        lt.Cmd('\Q');
        % 關閉 向前模擬
        lt.Cmd('\O"ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last].FORWARD_SIM_FUNCTION[Forward Simulation]"');
        lt.Cmd('Enabled=No ');
        lt.Cmd('"Has Illuminance"=No ');
        lt.Cmd('"Has Intensity"=No ');
        lt.Cmd('\Q');
        % 向後模擬 + 角度輝度 設定
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
    cprintf('err',"完成\n")
else % (不刪除接收器)
    if ~any(NumReceiverArray == 14) % 13點模式
        ReceiverCheck(lt)
    else
        error("第14點模式不支援接收器保留, 請將buildRec設定為0如果要進行第14點模式")
    end
end
break
end
%% Begin Simulation
while (1)
cprintf('key',"LT 模擬中......")
tStart = tic;
lt.Cmd("\V3D");
lt.Cmd("BeginAllSimulations");
tEnd = toc(tStart);
cprintf('err',"完成\n")
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
%% 開啟向後模擬-角度輝度 (optional)
% lt.Cmd("\V3D");
% lt.Cmd(['LumViewAngularLuminanceChart "R13 向後_模擬"']);
%% 存取資料
while (1)
if saveLTRawData==1
    cprintf('key',"寫入資料中......")
    for N = NumReceiverArray
        meshKey = strcat("SURFACE_RECEIVER[R",num2str(Wr(N)),"].BACKWARD_ANGULAR_LUMINANCE[Angular Luminance].ANGULAR_LUMINANCE_MESH[Angular Luminance Mesh]");
        Xdim = lt.DbGet(meshKey, "X_Dimension");
        Ydim = lt.DbGet(meshKey, "Y_Dimension");
        dataArray=zeros(Ydim,Xdim); 
        [~,dataArray] = lt.GetMeshData(meshKey, dataArray(), "CellValue");
        dataArray=rot90(double(dataArray)); % 必要處理
        % 軸值
        xAxisFirst = str2double(string(lt.DbGet(meshKey,"XCellCenterAt",def,1)));
        xAxisFinal = str2double(string(lt.DbGet(meshKey,"XCellCenterAt",def,Xdim)));
        xAxisArray = linspace(xAxisFirst,xAxisFinal,Xdim);
        yAxisFirst = str2double(string(lt.DbGet(meshKey,"YCellCenterAt",def,1)));
        yAxisFinal = str2double(string(lt.DbGet(meshKey,"YCellCenterAt",def,Ydim)));
        yAxisArray = linspace(yAxisFirst,yAxisFinal,Ydim)';
        % Raw Data 寫入 Excel
        strAtA1 = {'V\L'};
        writecell(strAtA1,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','A1','AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','B1','AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(yAxisArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','A2','AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(dataArray,excelFilepath,'Sheet',strcat("R",num2str(Wr(N))),'Range','B2','AutoFitWidth',false); % 搭配 LT 寫出格試% LT Fig (png) 存檔
        pngstr = strcat("R",num2str(Wr(N)),".png");
        pngFullFile = fullfile(fullDirc,pngstr);
        lt.Cmd("\V3D");
        lt.Cmd(strcat('LumViewAngularLuminanceChart "R',num2str(Wr(N)),' 向後_模擬"'));
        lt.Cmd("CopyToClipboard");
        lt.SetOption('ShowFileDialogBox', 0);
        lt.Cmd(strcat("PrintToFile """,pngFullFile,""""));
        lt.SetOption('ShowFileDialogBox', 1);
        lt.Cmd('Dismiss'); % 關閉視窗
    end
    cprintf('err',"完成\n")
end   
break
end
%% 記錄峰值誤差 (20220609 By GY)
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
        warning("有接收面沒有達至設定光線數")
    end
end
if errorCheck == 1
    disp(strcat("max(最大嘗試次數) = ",num2str(max(actualRayMultiplierArray(:)))));
else
    disp("所有接收面達至設定光線數.")
end
disp("峰值誤差")
disp(errorPeakPercentageArray)
break
end
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
function checkExcelFileName(fullfilename)
    numtest = strlength(fullfilename);
    if numtest > 218
        error("存出 Excel 完整檔名字元數必須小於 218")
    end
end