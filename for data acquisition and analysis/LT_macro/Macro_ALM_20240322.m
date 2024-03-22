%% 向後角輝 巨集 %%
% Author: GY (project since 20220614) (Ref by BM)
% Last Update: 20240322
% content: first version
close all;clear;clc
tStartFromBegining = tic;
%% 參數設定
% options
LTID = 13896;               % 同時開多個LT，指定用
saveLTFile = 0;             % 存取 LT File 開關
saveLTRaw = 1;              % 存取 LT excel, Fig.開關

% root folder: will create folder here
rootFolder = "";                                % 母資料夾決定. "": choose folder; cd: current folder (不用引號); "abc/def...": other path
    customLineFolder = '0322';                      % 結果資料夾額外名稱: "MRM5 LT還原 Custom..."
    customLineFile = '';                        % 結果檔案額外名稱: "(LT還原) II名 Custom..."

customDirectionGridApodizer = 0;                % 光源角度分佈 (0:Lambertian，1:使用者自訂，-1:維持現狀不詢問)
customSurfaceGridApodizer = 1;                  % 光源空間分佈 (0:均勻，1:使用者自訂，-1:維持現狀不詢問)
    % 切趾檔參數
    TXTSwitch = 0;                              % 切趾檔開關,啟用請輸入下方三列參數
        Zero2One = 1;                           % 是否將 Matrix 中的 0 轉為 1
        HorSize = 48;                          % (mm)
        VerSize = 27;                           % (mm)
seed = 1;                                       % 光追種子碼，可為任意正整數 (default = 1)
        
% 向後角輝 Receiver 參數 (note: follow "cono" 格式)
smoothFactor = 0;       % 是否打開平滑. 0: 關閉, other: 3,5,7,...21 (限奇數)
moduleTop = 0.174;      % mm (接收面會自行 = 0.001mm)
xOffset = 0;            % 
yOffset = 0;            % 
receiverSizeHor = 5;    % mm (Also affect 14th Receiver)
receiverSizeVer = 5;    % mm (Also affect 14th Receiver)
ALMHalfSize = 2.5;      % mm 計量器大小半徑 (建議 < 接收器大小/2)
LgridNum = 360;         % (預設) 俯視: 經度 (0-360)
VgridNum = 90;          % (預設) 俯視: 緯度 (上半球 0-90)
expectedERR = 0.08;     % 預期峰值誤差
MRM = 20;               % 最大嘗試次數 (vs Actual Ray Multiplier ARM)
rayFactor = 10;         % 光線數額外加乘
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% 參數確認
while (1)
% 其他參數
dateString = datestr(now,'mm-dd-yyyy HH-MM');

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
% 平滑參數 check
smoothFactorPool = [0,3,5,7,9,11,13,15,17,19,21];
if ~any(smoothFactor==smoothFactorPool)
    error("平滑參數設定有誤")
end
% extra_string
if ~isequal(customLineFile,"")
    customLineFile = strcat("_",customLineFile);
end
break
end
%% 創建資料夾
while (1)
if ~isempty(customLineFolder)
    customLineFolder = strcat("_",customLineFolder);
end
targetDirc = strcat("LT 向後角輝 ",customLineFolder);  %自定義資料夾名稱
rootFolder = string(rootFolder);
if isempty(rootFolder) || isequal(rootFolder,"")
    rootFolder = uigetdir("","選擇目標資料夾 (將在該資料夾中建立 LT 向後角輝資料夾)");
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

%% 連結 LT (Citrix Version)
while (1)
def=System.Reflection.Missing.Value;
% ltcom64path='C:\Program Files\Optical Research Associates\LightTools 9.1.1\Utilities.NET\LTCOM64.dll';      %LTCOM64.dll路徑
ltcom64path='C:\Program Files (x86)\Common Files\Optical Research Associates\LightTools\LTCOM64.dll';      %LTCOM64.dll路徑_Critrix!!
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

%% seed info
seed_firstTime = 1;
disp(strcat("[info]: 光追種子碼: ",num2str(seed)))

%% 峰值誤差與光線數估計
while (1)
ERR = expectedERR;
rayNum = (LgridNum * VgridNum)/ERR^2;
rayNum = rayNum * rayFactor;
break
end

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
    excelFilepath = fullfile(fullDirc,strcat("(LT 角輝) ",name,customLineFile,".xlsx"));  % 存結果 位置
    checkExcelFileName(excelFilepath);
    imageFilepath = fullfile(fullDirc,strcat("(LT 角輝) ",name,customLineFile,".png"));   % 存結果 位置
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

    %% Receiver Build
    while (1)
    cprintf('key',"LT 模型建立中......")
    ReceiverDelete(lt);     % 刪除當前所有 Dummy Plane
    sourceName = LightSourceSetup(lt);   % 空間網格設為均勻
    
    lt.Cmd("\V3D"); % 切到世界座標介面
    % dummy plane 建立
    lt.Cmd('DummyPlane ');
    lt.Cmd('XYZ'); % 虛擬面原點
    lt.Cmd(strcat(num2str(xOffset),',',num2str(yOffset),',',num2str(moduleTop+0.001)));
    lt.Cmd('XYZ'); % 虛擬面法線點 (與原點相減為法線)
    lt.Cmd(strcat(num2str(xOffset),',',num2str(yOffset),',',num2str(moduleTop+0.001+1)));
    lt.Cmd('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"');
    lt.Cmd(strcat('Name=D'));
    lt.Cmd('\Q');
    lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Width",receiverSizeHor);
    lt.DbSet("COMPONENTS[Components].PLANE_DUMMY_SURFACE[D]","Height",receiverSizeVer);
    % 接收器 建立
    lt.Cmd(strcat('\O"LENS_MANAGER[1].COMPONENTS[Components].PLANE_DUMMY_SURFACE[@Last]"'));
    lt.Cmd('"Add Receiver"=');
    lt.Cmd('\Q');
    lt.Cmd('\O"LENS_MANAGER[1].ILLUM_MANAGER[Illumination Manager].RECEIVERS[Receiver List].SURFACE_RECEIVER[@Last]"');
    lt.Cmd(strcat('Name=R'));
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
    
    cprintf('err',"完成\n")
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
    %% 檢查 嘗試次數
    errActual = lt.DbGet("SURFACE_RECEIVER[R].BACKWARD_ANGULAR_LUMINANCE[Angular_Luminance].ANGULAR_LUMINANCE_MESH[Angular_Luminance_Mesh]","ErrorAtPeak_Percent");
    cprintf('text',"峰值誤差: %.2f %% \n",round(errActual*100)/100)
    ARM = lt.DbGet("SURFACE_RECEIVER[R].BACKWARD_ANGULAR_LUMINANCE[Angular_Luminance]","Actual_Ray_Multiplier");
    if MRM < ARM;warning("最大光線嘗試次數 小於 實際光線嘗試次數");end
    
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
    if saveLTRaw==1
        cprintf('key',"寫入資料中......")
        meshKey = strcat("SURFACE_RECEIVER[R].BACKWARD_ANGULAR_LUMINANCE[Angular Luminance].ANGULAR_LUMINANCE_MESH[Angular Luminance Mesh]");
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
        writecell(strAtA1,excelFilepath,'Sheet',"R",'Range','A1','AutoFitWidth',false);
        writematrix(xAxisArray,excelFilepath,'Sheet',"R",'Range','B1','AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(yAxisArray,excelFilepath,'Sheet',"R",'Range','A2','AutoFitWidth',false); % 搭配 LT 寫出格試
        writematrix(dataArray,excelFilepath,'Sheet',"R",'Range','B2','AutoFitWidth',false); % 搭配 LT 寫出格試
        % LT Fig (png) 存檔
        lt.Cmd("\V3D");
        lt.Cmd(strcat('LumViewAngularLuminanceChart "R 向後_模擬"'));
        lt.Cmd("CopyToClipboard");
        lt.SetOption('ShowFileDialogBox', 0);
        lt.Cmd(strcat("PrintToFile """,imageFilepath,""""));
        lt.SetOption('ShowFileDialogBox', 1);
        lt.Cmd('Dismiss'); % 關閉視窗
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

    % 最後清掉 object
    lt.ListDelete(ObjectList);
end
function checkExcelFileName(fullfilename)
    numtest = strlength(fullfilename);
    if numtest > 218
        error("存出 Excel 完整檔名字元數必須小於 218")
    end
end