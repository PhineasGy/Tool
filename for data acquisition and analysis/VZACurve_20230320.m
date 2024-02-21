% VZA Curve Coding (Idea From Louie's Tidy Summary)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
clc;clear;close all
%% Current Version: 
% 20230220: first version
% 20230223: add Bold,Italic,"Calibri", add data for Bubble Plot
% 20230315: bug fix, VD length, no data(nan-->-777)
% 20230320: rangeMax setting
tic
try
%% 使用者輸入
% 讀取 excel
% excel 內容 以 VZA excel 輸出格式為主
excelName = "test.xlsx";                        % 輸入 ""/''/[]: 選取檔案 or 輸入 "....xlsx"
sheetNameMode = 1;                              % 0: 手動輸入需要的分頁名稱 , 1: 自動抓 Excel 所有分頁
    sheetNameArray = ["P1-33553","P0.9-33553","P0.8-33553"]; % 手動輸入分頁名稱
    %% other example with sheetNameArray (for loop) %%
    % numArray = [1,0.9,0.8];
    % sheetNameArray = strings(1,3);
    % for ss = 1:3
    %     sheetNameArray(ss) = strcat("P",num2str(numArray(ss)),"-33553"); % 手動輸入分頁名稱
    % end
    %% Filter Mode 去除或包含 sheet. Input [] to ignore.
    filterMode = 1;
    filterExcludeString = ["Final Summary"];    % make sure using string type "" (char type '' is not allowed)
    filterExcludeIndex = [];                    % integer only
    filterIncludeString = [];                   % make sure using string type "" (char type '' is not allowed)
    filterIncludeIndex = [];                    % integer only

% 其他設定
writeExcel = 1;
    sheetNameNew = "Final Summary";
plotCurve = 1;
    plotMinusPlus = 1;
    plotCenterRange = 1;
writeCurve = 1;
    figureFilename = "P1 system.png";           % 輸出範例: "(CR) ___.png" (CR for Center/Range)(MP For Minus/Plus)
markerDifferentShape = 1;                       % marker shape 是否不同
rangeIndex = 0;                                 % 紀錄 range table ==> 0: HVA, 1:VVA, 2: HVA+VVA
VVAOnlyorHVAOnly = 0;                           % set to 1 to ignore "nan" value error
WDArray = [400,500,600,700];                    % 每個分頁 WD 陣列
OOVA = [30,0];                                  % OOVA: [VVA,HVA]
rangeMaxHVA = 40;
rangeMaxVVA = 40;
%% Processing
cprintf(strcat("Process Start.\n"))
%% 抓 Excel 檔名
if isempty(excelName) || isequal(excelName,"")
    excelPathname = [];
    [excelFilename, excelPathname] = uigetfile({strcat(excelPathname,'*.xlsx')}, 'LimitedVA Excel');
    if ~ischar(excelPathname) 
    return;end
else
    excelPathname = excelName;
end
%% Sheet 處理 (with filter)
if sheetNameMode == 1
    allSheets = sheetnames(excelPathname);
elseif sheetNameMode == 0
    allSheets = sheetNameArray;
end
sheetCandidate = allSheets;
if size(sheetCandidate,1) > 1;sheetCandidate = sheetCandidate';end
if filterMode == 1
    [sheetCandidate,errCheck] = SheetnameFilter(sheetCandidate,...
    filterExcludeString,filterExcludeIndex,filterIncludeString,filterIncludeIndex);
    if errCheck == 1;return;end
    cprintf(strcat("sheetname (after filter): ",strjoin(sheetCandidate,', '),"\n"))
elseif filterMode == 0
    cprintf(strcat("sheetname: ",strjoin(sheetCandidate,', '),"\n"))
end

%% RangeTable 
% 可選擇 存 HVA or VVA or VVA+HVA
szRange = [length(sheetCandidate),length(WDArray)+2];
varTypesRange = ["string",repmat("double",[1,length(WDArray)+1])];
WDStr = strings(1,length(WDArray));
for wd = 1: length(WDArray)
    WDStr(wd) = strcat("WD",num2str(WDArray(wd)));
end
varNamesRange = ["Term",WDStr,"Sum"];
rangeTable = table('Size',szRange,'VariableTypes',varTypesRange,'VariableNames',varNamesRange);

%% 作圖預處理
figMP = figure;
figCR = figure;
if markerDifferentShape == 1
    % All shape: ["-o","-+","-*","-.","-x","-_","-|","-square","-diamond","-^","-v","->","-<","-pentagram","-hexagram"];
    markerArray = ["-v","-^","->","-<","-o","-square","-diamond","-pentagram","-hexagram","-x","-*","-+","-.","-_","-|"];
elseif markerDifferentShape == 0
    markerArray = repmat("-v",[1,length(sheetCandidate)]);
end
%% Sheet Loop
sheetCount = 0;
VATable = cell(1,length(sheetCandidate));
dataCell = cell(1,length(sheetCandidate));
for whichSheet = sheetCandidate
    sheetCount = sheetCount + 1;
    %%  讀取 Excel VZA 資料
    if isempty(excelName) || isequal(excelName,"")
        VAData = readtable(fullfile(excelPathname, excelFilename),'Sheet',whichSheet);
    else
        VAData = readtable(excelName,'Sheet',whichSheet);
    end
    VATable{sheetCount} = VAData;
    %% table check
    % Excel Data:
    % 1. row 1, row 2 不可含有數值 (字串可以)
    % 2. row 3 開始是 Limited VA 數值起點
    WDTemp = 1:10:1+10*(length(WDArray)-1);
    for WDCheck = 1:length(WDArray)
        temp = table2array(VAData(WDTemp(WDCheck):WDTemp(WDCheck)+3,2:3));
        if ~isnumeric(temp)
            cprintf('err',"[error] 解析資料錯誤 (system stopped)\n")
            cprintf("Excel Data:\n")
            cprintf("1. row 1, row 2 不可含有數值 (字串可以)\n")
            cprintf("2. row 3 開始是 Limited VA 數值起點\n")
            beep
            return
        elseif any(isnan(temp))
            if VVAOnlyorHVAOnly == 0
                cprintf('err',"[error] 解析資料錯誤 (system stopped)\n")
                cprintf("Excel Data:\n")
                cprintf("1. row 1, row 2 不可含有數值 (字串可以)\n")
                cprintf("2. row 3 開始是 Limited VA 數值起點\n")
                beep
                return
            elseif VVAOnlyorHVAOnly == 1 % 只跑 VVA or HVA 時， 將另一組 nan 值改為 -777
                temp(isnan(temp)) = -777;
            end
        end
    end
    %% 存值
    tempValue = nan(4,length(WDArray));
    WDCount = 0;
    for whichWD = 1:length(WDArray)
        WDCount = WDCount + 1;
        % Louie method: mean(pass,fail)
        tempValue(:,whichWD) = mean(table2array(VAData(1+10*(WDCount-1):1+10*(WDCount-1)+3,2:3)),2);
    end
    % -777 調整為 OOVA
    % 777 drop error
%     t1 = tempValue(1:2,:);t1(t1 == -777) = OOVA(2);
%     t2 = tempValue(3:4,:);t2(t2 == -777) = OOVA(1);
    t1 = tempValue(1:2,:);t1(t1 == -777) = nan;
    t2 = tempValue(3:4,:);t2(t2 == -777) = nan;
    tempValue = [t1;t2];
    if any(tempValue==777)
        cprintf('err',"[error] 777 detected. (systemo stopped)")
        beep
        return
    end
    % value each WD
    HVAMinus_WDArray = tempValue(1,:);
    HVAPlus_WDArray = tempValue(2,:);
    HVACenter_WDArray = 0.5*(HVAPlus_WDArray + HVAMinus_WDArray);
    HVARange_WDArray = HVAPlus_WDArray - HVAMinus_WDArray; 
    VVAMinus_WDArray = tempValue(3,:);
    VVAPlus_WDArray = tempValue(4,:);
    VVACenter_WDArray = 0.5*(VVAPlus_WDArray + VVAMinus_WDArray);
    VVARange_WDArray = VVAPlus_WDArray - VVAMinus_WDArray;
    
    %% 作圖
    % 兩張圖: (minus/plus + center/range)
    if plotCurve == 1
        if plotMinusPlus == 1
            figure(figMP);
            axisMPArray = cell(1,4);
            VAPlotArray = {HVAPlus_WDArray,VVAPlus_WDArray,HVAMinus_WDArray,VVAMinus_WDArray};
            for pp = 1:4
                axisInterest = subplot(2,2,pp);
                axisMPArray{pp} = axisInterest;
                hold(axisInterest,'on')
                grid(axisInterest,'on')
                plot(axisInterest,WDArray, VAPlotArray{pp},markerArray(sheetCount),'MarkerFaceColor','auto',...
                    'LineWidth',2,'MarkerSize',8);
            end
        end
        if plotCenterRange == 1
            figure(figCR);
            axisCRArray = cell(1,4);
            VAPlotArray = {HVACenter_WDArray,VVACenter_WDArray,HVARange_WDArray,VVARange_WDArray};
            for pp = 1:4
                axisInterest = subplot(2,2,pp);
                axisCRArray{pp} = axisInterest;
                hold(axisInterest,'on')
                grid(axisInterest,'on')
                plot(axisInterest,WDArray, VAPlotArray{pp},markerArray(sheetCount),'MarkerFaceColor','auto',...
                    'LineWidth',2,'MarkerSize',8);
            end
        end
    end
    %% Data 整理 (For Excel)
    HVACenterAvg = mean(HVACenter_WDArray(~isnan(HVACenter_WDArray)));
    HVARangeSum = sum(HVARange_WDArray(~isnan(HVARange_WDArray)));
    VVACenterAvg = mean(VVACenter_WDArray(~isnan(VVACenter_WDArray)));
    VVARangeSum = sum(VVARange_WDArray(~isnan(VVARange_WDArray)));
   
    HVACRAvgSum_str = strjoin([string(round(HVACenterAvg)),string(round(HVARangeSum))],", ");
    VVACRAvgSum_str = strjoin([string(round(VVACenterAvg)),string(round(VVARangeSum))],", ");

    HVAMinus_str = string(round(HVAMinus_WDArray));
    HVAMinus_str(ismissing(HVAMinus_str)) = "nan";
    VVAMinus_str = string(round(VVAMinus_WDArray));
    VVAMinus_str(ismissing(VVAMinus_str)) = "nan";
    HVAPlus_str = string(round(HVAPlus_WDArray));
    HVAPlus_str(ismissing(HVAPlus_str)) = "nan";
    VVAPlus_str = string(round(VVAPlus_WDArray));
    VVAPlus_str(ismissing(VVAPlus_str)) = "nan";

    HVACenter_str = string(round(HVACenter_WDArray));
    HVACenter_str(ismissing(HVACenter_str)) = "nan";
    VVACenter_str = string(round(VVACenter_WDArray));
    VVACenter_str(ismissing(VVACenter_str)) = "nan";
    HVARange_str = string(round(HVARange_WDArray));
    HVARange_str(ismissing(HVARange_str)) = "nan";
    VVARange_str = string(round(VVARange_WDArray));
    VVARange_str(ismissing(VVARange_str)) = "nan";
    
    WDLength = length(WDArray);
    HVAMP_str = strings(WDLength,1);VVAMP_str = strings(WDLength,1);
    HVACR_str = strings(WDLength,1);VVACR_str = strings(WDLength,1);
    for ww = 1:WDLength
        HVAMP_str(ww) = strjoin([HVAMinus_str(ww),HVAPlus_str(ww)],", ");
        VVAMP_str(ww) = strjoin([VVAMinus_str(ww),VVAPlus_str(ww)],", ");
        HVACR_str(ww) = strjoin([HVACenter_str(ww),HVARange_str(ww)],", ");
        VVACR_str(ww) = strjoin([VVACenter_str(ww),VVARange_str(ww)],", ");
    end
    dataCell{sheetCount} = {HVAMP_str,VVAMP_str,HVACR_str,VVACR_str,HVACRAvgSum_str,VVACRAvgSum_str};

    % rangeTable 存值
    rangeTable{sheetCount,"Term"} = whichSheet;
    switch rangeIndex % 0: HVA, 1:VVA, 2: HVA+VVA
        case 0
            tempRange = HVARange_WDArray;
            tempRange(isnan(HVARange_WDArray)) = 0;
            
        case 1
            tempRange = VVARange_WDArray;
            tempRange(isnan(VVARange_WDArray)) = 0;
        case 2
            tempRange1 = HVARange_WDArray;
            tempRange1(isnan(HVARange_WDArray)) = 0;
            tempRange2 = VVARange_WDArray;
            tempRange2(isnan(VVARange_WDArray)) = 0;
            tempRange = tempRange1 + tempRange2;
    end
    rangeTable{sheetCount,2:length(WDArray)+1} = tempRange;
    rangeTable{sheetCount,end} = sum(tempRange);
end
%% 作圖後處理
if plotCurve == 1
    if plotMinusPlus == 1
        yticksArray = {[0:5:30],[0:5:60],[-30:5:0],[0:5:60]};
        ylimArray = {[0 30],[0 60],[-30 0],[0 60]};
        titleArray = ["WD v.s HVA+","WD v.s VVA+","WD v.s HVA-","WD v.s VVA-"];
        for ww = 1:4       
            axisInterest = axisMPArray{ww};
            xticks(axisInterest,[WDArray(1)-100:50:WDArray(end)+150]) % fixed
            xlim(axisInterest,[WDArray(1)-100,WDArray(end)+150]) % fixed
            xlabel(axisInterest,"WD (mm)") % fixed
            yticks(axisInterest,yticksArray{ww})
            ylim(axisInterest,ylimArray{ww})
            ylabel(axisInterest,"Angle") % fixed
            title(axisInterest,titleArray(ww))
            legend(axisInterest,sheetCandidate,Box="off")  % fixed
            fontsize(axisInterest,16,"points")  % fixed
        end
    end
    if plotCenterRange == 1
        yticksArray = {[-20:5:20],[OOVA(1)-20:5:OOVA(1)+20],[0:5:rangeMaxHVA],[0:5:rangeMaxVVA]};
        ylimArray = {[-20 20],[OOVA(1)-20 OOVA(1)+20],[0 rangeMaxHVA],[0 rangeMaxVVA]};
        titleArray = ["WD v.s HVA Center","WD v.s VVA Center","WD v.s HVA Range","WD v.s VVA Range"];
        for ww = 1:4       
            axisInterest = axisCRArray{ww};
            xticks(axisInterest,[WDArray(1)-100:50:WDArray(end)+150]) % fixed
            xlim(axisInterest,[WDArray(1)-100,WDArray(end)+150]) % fixed
            xlabel(axisInterest,"WD (mm)") % fixed
            yticks(axisInterest,yticksArray{ww})
            ylim(axisInterest,ylimArray{ww})
            ylabel(axisInterest,"Angle") % fixed
            title(axisInterest,titleArray(ww))
            legend(axisInterest,sheetCandidate,Box="off")  % fixed
            fontsize(axisInterest,16,"points")  % fixed
        end
    end
    %% 放大Curve視窗
    figMP.WindowState = 'maximized';
    figCR.WindowState = 'maximized';
toc
end
%% Write Curve / Excel
if writeCurve == 1
    if plotMinusPlus == 1
        saveas(figMP,strcat("(MP) ",figureFilename));
    end
    if plotCenterRange == 1
        saveas(figCR,strcat("(CR) ",figureFilename));
    end    
end
%% Write Excel
while 1
if writeExcel == 1
    %% 整理資料格式
    tableCell = cell(1,4); % 1. HVA-,HVA+ 2. VVA-,VVA+ 3. HVACenter,HVARange 4. VVACenter,VVARange
    sz = [length(WDArray),length(sheetCandidate)+1];
    varTypes = ["double",repmat("string",[1,length(sheetCandidate)])];
    varNames = ["WD",sheetCandidate];
    tableCell = table('Size',sz,'VariableTypes',varTypes,'VariableNames',varNames);
    tableCell(:,1) = table(WDArray'); % WD Array
    HVAMPTable = tableCell; VVAMPTable = tableCell; 
    HVACRTable = tableCell; VVACRTable = tableCell;
    HVACRAvgSumFinal = ["Avg, Sum"];  VVACRAvgSumFinal = ["Avg, Sum"];
    HVACRAvgSumArray = strings(1,length(sheetCandidate));
    VVACRAvgSumArray = strings(1,length(sheetCandidate));
    % dataCell{whichSheet} = {HVAMP_str,VVAMP_str,HVACR_str,VVACR_str,HVACRAvgSum_str,VVACRAvgSum_str};
    for ss = 1:length(sheetCandidate)
        temp = dataCell{ss};
        HVAMPTable(:,1+ss) = table(temp{1});
        VVAMPTable(:,1+ss) = table(temp{2});
        HVACRTable(:,1+ss) = table(temp{3});
        VVACRTable(:,1+ss) = table(temp{4});
        HVACRAvgSumArray(ss) = temp{5};
        VVACRAvgSumArray(ss) = temp{6};
    end
    HVACRAvgSumFinal = [HVACRAvgSumFinal,HVACRAvgSumArray]; %#ok<AGROW> 
    VVACRAvgSumFinal = [VVACRAvgSumFinal,VVACRAvgSumArray]; %#ok<AGROW> 
    %% Write Excel
    % position: 
    % title --> C1 (合併 D1 E1...), G1 (合併 H1 I1), C8 (合併 D8 E8), G8 (合併 H8 I8)
    % table --> B2, F2, B9, F9
    % CRAvgSum --> B14, F14
    % sheet 數量決定 下一個 column 位置
    % first Column/Row: title C1, table B2
    cprintf("Writing Excel ......");
    nextColOffset = 1+length(sheetCandidate); % EX: 4
    nextColIndexStr_Title = xlscol(xlscol('C')+nextColOffset); % EX: G
    nextColIndexStr_Table = xlscol(xlscol('B')+nextColOffset); % EX: F
    nextRowOffset = 3+length(WDArray); % EX: 7
    nextRowIndexStr_Title = string(1+nextRowOffset); % EX: 8
    nextRowIndexStr_Table = string(2+nextRowOffset); % EX: 9
    rightdownCornerInd = [(xlscol('B')+nextColOffset*2-1),nextRowOffset*2];

    titleString_Excel = ["HVA-, HVA+","VVA-, VVA+","HVA Center, HVA Range","VVA Center, VVA Range"];
    titleIndex_Excel = ["C1",strcat(nextColIndexStr_Title,"1"),...
                        strcat("C",nextRowIndexStr_Title),strcat(nextColIndexStr_Title,nextRowIndexStr_Title)];
    table_Excel = {HVAMPTable,VVAMPTable,HVACRTable,VVACRTable};
    tableIndex_Excel = ["B2",strcat(nextColIndexStr_Table,"2"),...
                        strcat("B",nextRowIndexStr_Table),strcat(nextColIndexStr_Table,nextRowIndexStr_Table)];
    HVAAvgSumIndex_Excel = strcat("B",string(nextRowOffset*2));
    VVAAvgSumIndex_Excel = strcat(nextColIndexStr_Table,string(nextRowOffset*2));
    DataStartIndex = ["B1",strcat(xlscol(xlscol('B')+nextColOffset*2-1),string(nextRowOffset*2))];
    for tt = 1:4
        writematrix(titleString_Excel(tt),excelPathname,'Sheet',sheetNameNew,'Range',titleIndex_Excel(tt));
        writetable(table_Excel{tt},excelPathname,'Sheet',sheetNameNew,'Range',tableIndex_Excel(tt));
    end
    writematrix(HVACRAvgSumFinal,excelPathname,'Sheet',sheetNameNew,'Range',HVAAvgSumIndex_Excel)
    writematrix(VVACRAvgSumFinal,excelPathname,'Sheet',sheetNameNew,'Range',VVAAvgSumIndex_Excel)
    %% Excel 後處理
    % 自動調整 寬
    hExcel = actxserver('Excel.Application');
    % auto open excel
%     hExcel.Visible = 1;
    % open workbook
    hWorkbook = hExcel.Workbooks.Open(strcat(cd,"\",excelPathname));
    % locate sheet index of "Final Summary"
    sheetNum = hWorkbook.Sheets.Count; 
    FS_sheetIndex = 0;
    for ex = 1:sheetNum
        if matches(hWorkbook.Sheets.Item(ex).Name,sheetNameNew)
            FS_sheetIndex = ex;
            break
        end
    end
    % merge cell
    rangeInd1 = strcat("C1:",xlscol(xlscol('C')+length(sheetCandidate)-1),"1");
    rangeInd2 = strcat(nextColIndexStr_Title,"1:",xlscol(xlscol(nextColIndexStr_Title)+length(sheetCandidate)-1),"1");
    rangeInd3 = strcat("C",string(nextRowIndexStr_Title),":",xlscol(xlscol('C')+length(sheetCandidate)-1),string(nextRowIndexStr_Title));
    rangeInd4 = strcat(nextColIndexStr_Title,string(nextRowIndexStr_Title),":",xlscol(xlscol(nextColIndexStr_Title)+length(sheetCandidate)-1),string(nextRowIndexStr_Title));
    FS_sheet = hWorkbook.Sheets.Item(FS_sheetIndex);
    rangeTemp = FS_sheet.Range(rangeInd1);
    rangeTemp.MergeCells = true;
    rangeTemp = FS_sheet.Range(rangeInd2);
    rangeTemp.MergeCells = true;
    rangeTemp = FS_sheet.Range(rangeInd3);
    rangeTemp.MergeCells = true;
    rangeTemp = FS_sheet.Range(rangeInd4);
    rangeTemp.MergeCells = true;
    % alignment
    FS_sheet.Cells.EntireColumn.AutoFit;
    FS_sheet.Cells.EntireRow.AutoFit;
    FS_sheet.Cells.HorizontalAlignment = 3;
    FS_sheet.Cells.VerticalAlignment = 2;
    % border (slow)
    for xx = 2:rightdownCornerInd(1)
        for yy = 1:rightdownCornerInd(2)
            colStr = xlscol(xx);
            rowstr = string(yy);
            rangeStr = strcat(colStr,rowstr);
            selectedRange = FS_sheet.Range(rangeStr);
            selectedRange.BorderAround(1,2);
        end
    end
    % font (final row: bold and italic)
    rangeTemp = FS_sheet.Range(strcat(HVAAvgSumIndex_Excel,":",DataStartIndex(2)));
    rangeTemp.Font.Bold = 1;
    rangeTemp.Font.Italic = 1;
    
    % select
    FS_sheet.Select;
    FS_sheet.Cells.Font.Name = 'Calibri';
    FS_sheet.Range("A1").Select;
    hWorkbook.Save
    hWorkbook.Close
    hExcel.Quit
    cprintf("Done.\n");
end
break % while 1
end
cprintf("Process Complete.\n");
catch e
    beep
    cprintf("\n")
    for eEach = length(e.stack):-1:1
        disp("Location");
        disp(e.stack(eEach));
        matlab.desktop.editor.openAndGoToLine(e.stack(eEach).file,e.stack(eEach).line);
    end
    disp('There was an error!');
    disp(e.message);
    return
end
%% function
function [sheetCandidate,errCheck] = SheetnameFilter(sheetCandidate,...
    filterExcludeString,filterExcludeIndex,filterIncludeString,filterIncludeIndex)
    errCheck = 0;
    % check type
    if ~(isempty(filterExcludeString) || isstring(filterExcludeString)) || ...
            ~(isempty(filterIncludeString) || isstring(filterIncludeString)) ||...
            ~(isempty(filterExcludeIndex) || isnumeric(filterExcludeIndex)) ||...
            ~(isempty(filterIncludeIndex) || isnumeric(filterIncludeIndex))
        cprintf('err',"[error] Wrong type for filter parameter detected. (system stopped)\n")
        beep
        errCheck = 1;
        return
    end
    
    % check index error 
    if any(filterExcludeIndex > length(sheetCandidate)) ||...
            any(filterIncludeIndex > length(sheetCandidate))
        cprintf('err',"[error] Maximum of filter index is larger than the number of sheets. (system stopped)\n")
        beep
        errCheck = 1;
        return
    end

    % check inlcude string should be all inside sheetCandidate
    try
        if ~all(matches(filterIncludeString,sheetCandidate))
            cprintf('err',"[error] unknown string detected in filterIncludeString (should be the string among sheetnames). (system stopped)\n")
            beep
            errCheck = 1;
            return
        end
    catch
    end

    % index transform to string
    if ~isempty(filterExcludeIndex)
        filterExcludeIndex = sheetCandidate(filterExcludeIndex);
    end
    if ~isempty(filterIncludeIndex)
        filterIncludeIndex = sheetCandidate(filterIncludeIndex);
    end

    % cell build (ignore empty term)
    filterCell = {filterExcludeString,filterExcludeIndex,filterIncludeString,filterIncludeIndex};
    filterNumArray = [1,2,3,4];
    filterNumArray(cellfun(@isempty,filterCell)) = [];
    filterCell(cellfun(@isempty,filterCell)) = [];
    
    % check in/exclude conflict error
    for whichfilter = length(filterCell)
        filterNow = filterCell{whichfilter};
        filterNum = filterNumArray(whichfilter);
        % check in/exclude conflict error
        dropConflictErr = 0;
        if filterNum == 1 || filterNum == 2 % compare 3 , 4 if exist
            try
                if any(matches(filterNow,filterCell{filterNumArray==3}))
                    dropConflictErr = 1;
                end
            catch
            end
            try
                if any(matches(filterNow,filterCell{filterNumArray==4}))
                    dropConflictErr = 1;
                end
            catch
            end
        elseif filterNum == 3 || filterNum == 4 % compare 1 , 2 if exist
            try
                if any(matches(filterNow,filterCell{filterNumArray==1}))
                    dropConflictErr = 1;
                end
            catch
            end
            try
                if any(matches(filterNow,filterCell{filterNumArray==2}))
                    dropConflictErr = 1;
                end
            catch
            end
        end
        if dropConflictErr == 1
            cprintf('err',"[error] same conflict value detected in 'FilterExclude' and 'FilterInclude' Index/String. (system stopped)\n")
            beep
            errCheck = 1;
            return
        end
    end

    % sheetname filter
    % sheetname : 排除掉 filterExcludeString(1) filterExcludeIndex(2) (||)
    % sheetname : 要包含 filterIncludeString(3) filterIncludeIndex(4) (&&)
    % filterCell
    % exlcude
    try
        sheetCandidate(matches(sheetCandidate,filterCell{filterNumArray==1})) = [];
    catch
    end
    try
        sheetCandidate(matches(sheetCandidate,filterCell{filterNumArray==2})) = [];
    catch
    end
    % include
    try
        sheetCandidate(~matches(sheetCandidate,filterCell{filterNumArray==3|filterNumArray==4})) = [];
    catch
    end
end
function b = xlscol(a)
%XLSCOL Convert Excel column letters to numbers or vice versa.
%   B = XLSCOL(A) takes input A, and converts to corresponding output B.
%   The input may be a number, a string, an array or matrix, an Excel
%   range, a cell, or a combination of each within a cell, including nested
%   cells and arrays. The output maintains the shape of the input and
%   attempts to "flatten" the cell to remove nesting.  Numbers and symbols
%   within strings or Excel ranges are ignored.
%
%   Examples
%   --------
%       xlscol(256)   % returns 'IV'
%
%       xlscol('IV')  % returns 256
%
%       xlscol([405 892])  % returns {'OO' 'AHH'}
%
%       xlscol('A1:IV65536')  % returns [1 256]
%
%       xlscol({8838 2430; 253 'XFD'}) % returns {'MAX' 'COL'; 'IS' 16384}
%
%       xlscol(xlscol({8838 2430; 253 'XFD'})) % returns same as input
%
%       b = xlscol({'A10' {'IV' 'ALL34:XFC66'} {'!@#$%^&*()'} '@#$' ...
%         {[2 3]} [5 7] 11})
%       % returns {1 [1x3 double] 'B' 'C' 'E' 'G' 'K'}
%       %   with b{2} = [256 1000 16383]
%
%   Notes
%   -----
%       CELLFUN and ARRAYFUN allow the program to recursively handle
%       multiple inputs.  An interesting side effect is that mixed input,
%       nested cells, and matrix shapes can be processed.
%
%   See also XLSREAD, XLSWRITE.
%
%   Version 1.1 - Kevin Crosby

% DATE      VER  NAME          DESCRIPTION
% 07-30-10  1.0  K. Crosby     First Release
% 08-02-10  1.1  K. Crosby     Vectorized loop for numerics.

% Contact: Kevin.L.Crosby@gmail.com

base = 26;
if iscell(a)
  b = cellfun(@xlscol, a, 'UniformOutput', false); % handles mixed case too
elseif ischar(a)
  if ~isempty(strfind(a, ':')) %#ok<STREMP> % i.e. if is a range
    b = cellfun(@xlscol, regexp(a, ':', 'split'));
  else % if isempty(strfind(a, ':')) % i.e. if not a range
    b = a(isletter(a));        % get rid of numbers and symbols
    if isempty(b)
      b = {[]};
    else % if ~isempty(a);
      b = double(upper(b)) - 64; % convert ASCII to number from 1 to 26
      n = length(b);             % number of characters
      b = b * base.^((n-1):-1:0)';
    end % if isempty(a)
  end % if ~isempty(strfind(a, ':')) % i.e. if is a range
elseif isnumeric(a) && numel(a) ~= 1
  b = arrayfun(@xlscol, a, 'UniformOutput', false);
else % if isnumeric(a) && numel(a) == 1
  n = ceil(log(a)/log(base));  % estimate number of digits
  d = cumsum(base.^(0:n+1));   % offset
  n = find(a >= d, 1, 'last'); % actual number of digits
  d = d(n:-1:1);               % reverse and shorten
  r = mod(floor((a-d)./base.^(n-1:-1:0)), base) + 1;  % modulus
  b = char(r+64);  % convert number to ASCII
end % if iscell(a)

% attempt to "flatten" cell, by removing nesting
if iscell(b) && (iscell([b{:}]) || isnumeric([b{:}]))
  b = [b{:}];
end % if iscell(b) && (iscell([b{:}]) || isnumeric([ba{:}]))
end