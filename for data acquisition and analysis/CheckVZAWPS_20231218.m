%% VZAWPS: Viewing Zone All-White Pupil Size 找 "裕度全白 PS" 工具
% ver: 20231218
% content: PS 擷取 num2str(PS,'%02d')
close all; 
clc;clear; 
%% User Input
% 量化 Method:
% 分子: summation of (maxPS Image - targetPS Image)
% 分母: LL 積分
% 看 PS 降到什麼程度之後 無法接受 (最大PS影像 - 每一張PS影像:檢查非零值數量)
% 請確保最大 PS 是全白的結果

autoMode = 1;
WD_PLarray = [400:100:700];
PSArray = [20,15:-1:8];                          % 請確保為遞減數列 EX: 15:-1:5

% 量化 method (分母)
humanFactor = 1;
    meshSize = 0;                           % only work when humanFactor = 0 (mm)

% L14 Portrait %
% LRA = 14;
% lensPitch = 1.991;
% panelSizeHor = 165.24;                      % panel horizontal size (mm)
% panelSizeVer = 293.76;                      % panel vertical size (mm)

% L20 Portrait %
LRA = 12;
lensPitch = 1.001;
panelSizeVer = 293.76;                      % panel horizontal size (mm)
panelSizeHor = 165.24;                      % panel vertical size (mm)

AWThreshold = 10;                           % unit: 千分比
showGLThreshold = 5;                        % 秀影像時二值化閥值

% 選圖規則: (mainly for software II)
% 請確保影像名稱包含: ..._WDR700_... 和 ..._PS15_... 的形式 (順序不定)
UserDefinedPattern = 0;
    WDRUserFront = "VD=0";
    WDRUserBack = ".00";
    WDRRule = 0;                            % 底線位置 --> 0: _WDR700_ , 1: _WDR700 , 2: WDR700_, 3: WDR700
    PSUserFront = "PR=";
    PSUserBack = ".00";
    PSRule = 0;                             % 底線位置 --> 0: _WDR700_ , 1: _WDR700 , 2: WDR700_, 3: WDR700
    PSnum9to09 = 1;
    
%% Process Start %%
%% check input
if any(diff(PSArray)>0)
    disp("[error]: 請確保 PSArray 為遞減數列 ")
    beep
    return
end
if autoMode == 0
    disp("[warning]: matlab 顯示影像解析度可能與真實情況有落差，請記得用其他開圖軟體 Double Check.")
end
%% 讀取檔案
while (1)
lengthFiles = 1;
if 1
    pathname=[];
    [filename, pathname] = uigetfile(strcat(pathname,'*.png'), '請選擇 LT 還原圖','MultiSelect','on');
    
    % 計算檔案數量
    if ~ischar(pathname) 
        return;end
    if ischar(filename)
        lengthFiles=1;
    else
        lengthFiles=length(filename);
    end
end
break
end

%% Lens Count
pn_left = -0.5 * panelSizeHor;
pn_right = 0.5 * panelSizeHor;
pn_up = -0.5 * panelSizeVer;
pn_down = 0.5 * panelSizeVer;
pn_array = [pn_left,pn_right,pn_up,pn_down];
rangeYPitch = lensPitch / cosd(LRA); % range_y Pitch
sizeforRangeY = panelSizeHor + panelSizeVer * abs(tand(LRA));
rangeY = - floor(0.5 * sizeforRangeY / rangeYPitch) : floor(0.5 * sizeforRangeY / rangeYPitch);
numLensYOriginal = length(rangeY);
%% WD PL Loop
WDCount = 0;
VZAWPS_Array = nan(1,length(WD_PLarray));
AWRatioArray = nan(1,length(WD_PLarray));
AWRatioTable = nan(length(PSArray),length(WD_PLarray));
AWRatioTable(1,:) = 0;
failRPNumTable = nan(length(PSArray),length(WD_PLarray));

imageArray = cell(length(PSArray),length(WD_PLarray));
for whichWD = WD_PLarray
    WDCount = WDCount + 1;
    WD = WD_PLarray(WDCount);
    disp("------------");
    cprintf('key',strcat("Current WD: ",num2str(WD),"\n"));
    %% read image  
    for pp = 1:length(PSArray)
        % 認 Pattern: _WDR700_... , ..._PS8_...
        PS = PSArray(pp);
        if UserDefinedPattern == 1
            % WDR
            switch WDRRule
                case 0
                    underline_before = "_";
                    underline_after = "_";
                case 1
                    underline_before = "_";
                    underline_after = "";
                case 2
                    underline_before = "";
                    underline_after = "_";
                case 3
                    underline_before = "";
                    underline_after = "";
            end
            ul_b = underline_before;
            WDRString = num2str(WD);
            ul_a = underline_after;
            WDRPattern = caseInsensitivePattern(strcat(ul_b,WDRUserFront,WDRString,WDRUserBack,ul_a));
            % PS
            switch PSRule
                case 0
                    underline_before = "_";
                    underline_after = "_";
                case 1
                    underline_before = "_";
                    underline_after = "";
                case 2
                    underline_before = "";
                    underline_after = "_";
                case 3
                    underline_before = "";
                    underline_after = "";
            end
            if PSnum9to09 == 1
                if any(PS==1:9)
                    add0 = "0";
                else
                    add0 = "";
                end
            else
                add0 = "";
            end
            ul_b = underline_before;
            PSString = num2str(PS);
            ul_a = underline_after;
            PSPattern = caseInsensitivePattern(strcat(ul_b,PSUserFront,add0,PSString,PSUserBack,ul_a));
        elseif UserDefinedPattern == 0
            WDRPattern = caseInsensitivePattern(strcat("_","WDR",num2str(WD),"_"));
            PSPattern = caseInsensitivePattern(strcat("_","PS",num2str(PS,'%02d'),"_"));
        end
        WDRContain = contains(filename,WDRPattern);
        PSContain = contains(filename,PSPattern);
        try
            imageString = filename{WDRContain & PSContain};
        catch
            beep;
            error("cannot extract WDR or PS value. (確保命名包含:... _WDR700_ ..._PS15_...)");
        end
        imageArray{pp,WDCount} = imread(fullfile(pathname,imageString));
    end
    maxPSImage = imageArray{1,WDCount};
    [imageSizeVer,imageSizeHor] = size(maxPSImage);
    pixelOffset = 10; % 可調
    avgGLFromMaxPSImage = mean(double(maxPSImage(round(imageSizeVer/2)-pixelOffset:round(imageSizeVer/2)+pixelOffset, ...
        round(imageSizeHor/2)-pixelOffset:round(imageSizeHor/2)+pixelOffset)),'all');
    %% 得量化分母 0508 
    % summation of lens Count and length
    if humanFactor == 1
        angRes = tand(1/120); % 人眼角分辨率
        meshSize = WD * angRes * 2;   
    end
    yCount = 0;
    sumTotal = 0;
    for whichY = rangeY
        yCount = yCount + 1;
        % intersection @ top and bottom panel
        if LRA == 0
            lensCenterTop = [-panelSizeVer*0.5 ; whichY*lensPitch];
            lensCenterBottom = [panelSizeVer*0.5 ; whichY*lensPitch];
        elseif LRA ~= 0
            Cpoint = [0; whichY * rangeYPitch]; % midpoint for each lens
            lensCenterTop = LensCenter_xy_Generator(-1,Cpoint,LRA,pn_array);
            lensCenterBottom = LensCenter_xy_Generator(1,Cpoint,LRA,pn_array);
        end
        % 取左線 (除了最後一根，取左右)
        temp1 = [0;-1*0.5*lensPitch;0];
        temp2 = rotz(LRA) * temp1;
        lensEdgeTopLeft = lensCenterTop + temp2(1:2); % [2D]
        lensEdgeBottomLeft = lensCenterBottom + temp2(1:2); % [2D]
        % length, summation
        lengthLL = norm(lensEdgeTopLeft-lensEdgeBottomLeft);
        sumTotal = sumTotal + (lengthLL/meshSize)*avgGLFromMaxPSImage;
        if yCount == numLensYOriginal
            temp1 = [0;1*0.5*lensPitch;0];
            temp2 = rotz(LRA) * temp1;
            lensEdgeTopRight = lensCenterTop + temp2(1:2); % [2D]
            lensEdgeBottomRight = lensCenterBottom + temp2(1:2); % [2D]
            lengthLL2 = norm(lensEdgeTopRight-lensEdgeBottomRight);
            sumTotal = sumTotal + (lengthLL2/meshSize)*avgGLFromMaxPSImage;
        end
    end
    %% 得量化分子
    % summation of difference between maxPS Image and targetPS Image
    VZAWPS = nan;
    AWRatio = nan;
    opts.Interpreter='tex';
    opts.Default = 'done!';
    for pp = 1:length(PSArray)
        currentImage = imageArray{pp,WDCount};
        diffImage = uint8(double(maxPSImage) - double(currentImage));
        %% Method 1 (old) 二值化找 Pixel 數
        failRPNumTable(pp,WDCount) = sum(imbinarize(diffImage)==1,'all');
        % 顯示用: 灰階以上為白
        diffImageForShow = diffImage;
        diffImageForShow(diffImageForShow>showGLThreshold) = 255;
        sumDiff = sum(double(diffImage),"all"); % 介在 0 - 1 之間
        AWRatio = sumDiff/sumTotal * 1000; % 千分比
        AWRatioTable(pp,WDCount) = AWRatio;
        if autoMode == 0
           F = figure;
                subplot(1,3,1), imshow(maxPSImage);
                    title(strcat(strcat("WD\color{red}",num2str(WD)),"\color{black} PS\color{blue}",num2str(PSArray(1))))
                subplot(1,3,2), imshow(currentImage);
                    title(strcat(strcat("WD\color{red}",num2str(WD)),"\color{black} PS\color{blue}",num2str(PSArray(pp))))                    
                subplot(1,3,3), imshow(diffImageForShow);
                    title(strcat(strcat("WD\color{red}",num2str(WD)),"\color{black} < PS\color{blue}",num2str(PSArray(1)),"\color{black} - PS \color{blue}",num2str(PSArray(pp)),"\color{black} > (二值化>",num2str(showGLThreshold),")"))
            F.WindowState = 'maximized';
            pause(1)
            prompt = {strcat("PS\color{blue} ",num2str(PSArray(1)),"\color{black} vs PS\color{blue} ",num2str(PSArray(pp))),...
                strcat("\color{black}all-white Ratio: ",num2str(AWRatio)," (千分比)")};
            answer = questdlg(prompt, ...
                    'Hell World', ...
                    'done!','next!','shut down',opts);
            switch answer
                case 'done!'
                    disp(strcat("min VZAWPS = ",num2str(PSArray(pp))))
                    disp(strcat("all-white Ratio: ",num2str(AWRatio)," (千分比)"))
                    VZAWPS = PSArray(pp);
                    VZAWPS_Array(WDCount) = PSArray(pp);
                    AWRatioArray(WDCount) = AWRatio;
                    break
                case 'next!'
                    close(F)
                    continue
                case 'shut down'
                    close(F)
                    disp("system stopped")
                    return    
            end 
        end
    end
    %% 跑每一張 PS 紀錄
    if autoMode == 1
        candidate = AWRatioTable(:,WDCount);
        failorPass = candidate > AWThreshold; % 0 0 "0" 1 1 1: 取雙引號位置 PS
        finalCheck = strfind(failorPass',[0 1]);
        if isempty(finalCheck)
            disp("無法確定 VZAWPS，PS Data 不足")
        elseif length(finalCheck) == 1
            VZAWPS_Array(WDCount) = PSArray(finalCheck);
            AWRatioArray(WDCount) = AWRatioTable(finalCheck,WDCount);
        else
            error("??")
        end
    end
    try close(F);catch;end
end % WD_PLArray
cprintf("---process complete---\n");
%% Table
T = table(WD_PLarray',VZAWPS_Array',AWRatioArray');
T.Properties.VariableNames = ["WD","VZAWPS","all-white Ratio"];
disp(T)
%% Total Data
method2 = fliplr(AWRatioTable');
disp("Total Data: 見 T2 (open T2)")
T2_1 = table(WD_PLarray');T2_1.Properties.VariableNames = ["WD"];
T2_2 = array2table(method2);T2_2.Properties.VariableNames = strcat("PS",string(fliplr(PSArray)));
T2 = [T2_1,T2_2];
open T2
%% Function
% LensCenter計算 (Rot_LRoV)
function LensCenter_xy = LensCenter_xy_Generator(x_updown,Cpoint,LRA,ll_array)
    if x_updown ==-1 % 找 X- 位置 up
        if LRA>=0 % 逆時針 斜線為左上到右下
            x_test = (cosd(LRA)/sind(LRA))*(ll_array(1)-Cpoint(2))+Cpoint(1);
            if x_test <= ll_array(3)
                x_test = ll_array(3);
                y_test =  (sind(LRA)/cosd(LRA))*(ll_array(3)-Cpoint(1))+Cpoint(2);
            elseif x_test > ll_array(3)
                y_test = ll_array(1);
            end
        elseif LRA<0 % 順時針 斜線為右上到左下
            x_test = (cosd(LRA)/sind(LRA))*(ll_array(2)-Cpoint(2))+Cpoint(1);
            if x_test <= ll_array(3)
                x_test = ll_array(3);
                y_test =  (sind(LRA)/cosd(LRA))*(ll_array(3)-Cpoint(1))+Cpoint(2);
            elseif x_test > ll_array(3)
                y_test = ll_array(2);
            end                
        end
    elseif x_updown==+1 % 找 X- 位置 down
        if LRA>=0 % 逆時針 斜線為左上到右下
            x_test = (cosd(LRA)/sind(LRA))*(ll_array(2)-Cpoint(2))+Cpoint(1);
            if x_test > ll_array(4)
                x_test = ll_array(4);
                y_test =  (sind(LRA)/cosd(LRA))*(ll_array(4)-Cpoint(1))+Cpoint(2);
            elseif x_test <= ll_array(4)
                y_test = ll_array(2);
            end
        elseif LRA<0 % 順時針 斜線為右上到左下
            x_test = (cosd(LRA)/sind(LRA))*(ll_array(1)-Cpoint(2))+Cpoint(1);
            if x_test > ll_array(4)
                x_test = ll_array(4);
                y_test =  (sind(LRA)/cosd(LRA))*(ll_array(4)-Cpoint(1))+Cpoint(2);
            elseif x_test <= ll_array(4)
                y_test = ll_array(1);
            end                
        end
    end
    LensCenter_xy = [x_test; y_test];
end