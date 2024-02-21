% VZA Alone Coding 
%% Current Version:
% 20230911: all 模式 table 參數型態修正: uint8 --> double
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
clc;clear;close all 
%% user input
% VA info setup
angleStep = 1;                          % must > 0
angleSweepVVA = [10 60];                   % absolute angle, must contain 2 value and (1) < (2).
angleSweepHVA = [-25 25];               % absolute angle, must contain 2 value and (1) < (2).
angleCenterVVA = 35;                    % absolute angle, stop point
angleCenterHVA = 0;                     % absolute angle, stop point
WD_list = [400 500 600 700];
% analysis setup
analyze_all = 1;                        % 0: 十字, 1: 全部分析
    % when 十字 mode:
    analyze_soft = 1;                       % II 圖類型: 0: hardware, 1: software
    HVA_VVA_mode = [1 2];                   % 1: HVA, 2: VVA, [1,2]: both
    criticalNumber = 1000;
    % write excel setup
    writeExcel = 0;                         % write limitedVA (PL loop support)
        excelFileName = "";                 % if "": "LimitedVA_....xlsx"
        excelSheetName = "";                % if "": sheetname = 1. (string or numeric support)
        dateStringOn = 0;                   % if off, make sure excel file is already closed

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% excel check
if writeExcel == 1
    if isequal(excelFileName,"")
        if dateStringOn == 1
            filename_excel = strcat("LimitedVA_",datestr(now,'mm-dd-yyyy HH-MM'),".xlsx");
        elseif dateStringOn == 0
            filename_excel = strcat("LimitedVA.xlsx");
        end
    else
        if dateStringOn == 1
            filename_excel = strcat(excelFileName,"_",datestr(now,'mm-dd-yyyy HH-MM'),".xlsx");
        elseif dateStringOn == 0
            filename_excel = strcat(excelFileName,".xlsx");
        end
    end
    if isequal(excelSheetName,"")
        excelSheetName = 1;
    end
    fullfilename = fullfile(cd,filename_excel);
    numtest = strlength(fullfilename);
    if numtest > 218
        disp(fullfilename)
        error("Excel 完整檔名過長")
    end
end
%% 選取 II 圖 
pathname = [];
[filename, pathname] = uigetfile({strcat(pathname,'*.png;',pathname,'*.bmp')}, '原圖','MultiSelect', 'on');
if ~ischar(pathname) 
    return;end
if ischar(filename)
    totalNumFile = 1;
else
    totalNumFile = length(filename);
end

%% 
if analyze_all == 1
    % 建立 complete table
    T = table('Size',[totalNumFile,3],'VariableTypes',["string","double","double"]);
    T.Properties.VariableNames = [{'filename'},{'fail number'},{'fail or pass'}];
    for whichII = 1:totalNumFile
        % step 1: 讀圖 + 二值化
        if totalNumFile == 1
            name = filename;
        else
            name = filename{whichFile};
        end
        filepath = fullfile(pathname, name);
        II = imbinarize(imread(filepath)); % [0,1]
        % step 2: VZA
        failPixelNumber = getFailNumber(II);
        % step 3: assign
        T(whichII,1) = {string(name)};
        T(whichII,2) = {failPixelNumber};
        T(whichII,3) = {failPixelNumber < criticalNumber};
    end
    disp("Analysis Process Done! (open T)")
    open T

elseif analyze_all == 0
    disp("VZA Process...(十字模式)");
    % 建立 VA 清單
    HVA_list = angleSweepHVA(1):angleStep:angleSweepHVA(end);
    VVA_list = angleSweepVVA(1):angleStep:angleSweepVVA(end);
    if any(HVA_VVA_mode==1)
        if HVA_list(end) ~= angleSweepHVA(end) 
            disp("[error] 請確保 angleSweepHVA(1):angleStep:angleSweepHVA(end) 包含..." + ...
            "angleSweepHVA(1) 和 angleSweepHVA(end) (system stopped)");
            beep
            return
        end
    else % HVA_VVA_mode == 2
        HVA_list = angleCenterHVA;
    end
    if any(HVA_VVA_mode==2)
        if VVA_list(end) ~= angleSweepVVA(end) 
            disp("[error] 請確保 angleSweepVVA(1):angleStep:angleSweepVVA(end) 包含..." + ...
            "angleSweepVVA(1) 和 angleSweepVVA(end) (system stopped)");
            beep
            return
        end
    else
        VVA_list = angleCenterVVA;
    end
    [HVAMesh,VVAMesh] = meshgrid(HVA_list,VVA_list);
    %% II 數量
    HVANum = 1; VVANum = 1; center_twice = 0; twice_str = "";
    if any(HVA_VVA_mode==1); HVANum = length(HVA_list);end
    if any(HVA_VVA_mode==2); VVANum = length(VVA_list);end
    if isequal(HVA_VVA_mode,[1,2]); center_twice = 1; twice_str = " (中心會計算到兩次)";end
    II_totalnumber = length(WD_list) * (HVANum + VVANum - 1);
    disp(strcat("II 數:",num2str(II_totalnumber),"+",num2str(center_twice*length(WD_list)),twice_str))

    %% VA Loop (十字版本)
    limitedVA = nan(2,4,length(WD_list));
    failNumberMatrix_WD = nan(length(VVA_list),length(HVA_list),length(WD_list));
    failPassMatrix_WD = nan(length(VVA_list),length(HVA_list),length(WD_list));
    numII = 0;
    for whichWD = 1:length(WD_list)
        WD = WD_list(whichWD);
        cprintf(strcat("-------------------------\n"))
        cprintf(strcat("處理中: WD",num2str(WD),"\n"))
        cprintf("II 圖分析: ")
        failNumberMatrix = nan(length(VVA_list),length(HVA_list));
        failPassMatrix = nan(length(VVA_list),length(HVA_list));
        for ii = HVA_VVA_mode % 1:HVA, 2:VVA
            switch ii
                case 1 % HVA loop
                    HVA_array = HVA_list;
                    VVA_array = angleCenterVVA;
                case 2
                    HVA_array = angleCenterHVA;
                    VVA_array = VVA_list;
            end
            for whichHVA = 1: length(HVA_array)
                for whichVVA = 1: length(VVA_array)
                    numII = numII + 1;
                    HVA = HVA_array(whichHVA);
                    VVA = VVA_array(whichVVA);
                    %% image 檔名解析
                    imageString = getImageStr(WD,VVA,HVA,filename,analyze_soft);
                    if contains(imageString,"TIR=True")
                        failNumberMatrix(HVAMesh==HVA & VVAMesh==VVA) = -1; % TIR
                        failPassMatrix(HVAMesh==HVA & VVAMesh==VVA) = -1; % TIR
                        cprintf(strcat(num2str(numII)," "));
                        if mod(numII,10)==0
                            cprintf('\n');
                        end
                        continue
                    end
                    %% 讀圖 + 二值化
                    II = imbinarize(imread(fullfile(pathname,imageString)));
                    %% step 2: VZA
                    failPixelNumber = getFailNumber(II);
                    %% step 3: assign
                    failNumberMatrix(HVAMesh==HVA & VVAMesh==VVA) = failPixelNumber;
                    failPassMatrix(HVAMesh==HVA & VVAMesh==VVA) = failPixelNumber < criticalNumber;
                    cprintf(strcat(num2str(numII)," "));
                    if mod(numII,10)==0
                        cprintf('\n');
                    end
                end
            end
        end
        cprintf("完成 \n");
        %% 建立 limited VA
        % limitedVA: [HVA- HVA+ VVA- VVA+] (row1: fail row2: pass)
        cprintf("建立 limited VA: ...")
        while 1 % HVA
            if any(HVA_VVA_mode==1)
                FP_OI = failPassMatrix(VVAMesh==angleCenterVVA);
                if size(FP_OI,1)==1;FP_OI = FP_OI';end
                % for both HVA- HVA+
                if all(FP_OI~=1) % 沒有任何位置是 Pass, 只有 TIR 或是 Fail
                    limitedVA(:,1:2,whichWD) = -777;
                    break
                end
                while 1 % for HVA- [limitedVA(:,1,:)]
                    if FP_OI(1) == 1 % all 1: 777
                        limitedVA(:,1,whichWD) = 777;
                        break
                    end
                    indTIR = min(strfind(FP_OI',[-1 1])); % TIR 後 Pass: 紀錄
                    if ~isempty(indTIR)
                        limitedVA(1,1,whichWD) = HVA_list(indTIR);    % fail
                        limitedVA(2,1,whichWD) = HVA_list(indTIR+1);  % pass  
                        break
                    end
                    ind1 = min(strfind(FP_OI',[0,1])); % 正常情形 0 --> 1
                    limitedVA(1,1,whichWD) = HVA_list(ind1);    % fail
                    limitedVA(2,1,whichWD) = HVA_list(ind1+1);  % pass                   
                    break
                end
                while 1 % for HVA+ [limitedVA(:,2,:)]
                    if FP_OI(end) == 1
                        limitedVA(:,2,whichWD) = 777;
                        break
                    end
                    indTIR = max(strfind(FP_OI',[1 -1])); % TIR 後 Pass: 紀錄
                    if ~isempty(indTIR)
                        limitedVA(1,2,whichWD) = HVA_list(indTIR+1);    % fail
                        limitedVA(2,2,whichWD) = HVA_list(indTIR);      % pass  
                        break
                    end
                    ind2 = max(strfind(FP_OI',[1,0]));
                    limitedVA(1,2,whichWD) = HVA_list(ind2+1);    % pass
                    limitedVA(2,2,whichWD) = HVA_list(ind2);  % fail
                break
                end               
            end
            break
        end
        while 1 % VVA
            if any(HVA_VVA_mode==2)
                FP_OI = failPassMatrix(HVAMesh==angleCenterHVA);
                if size(FP_OI,1)==1;FP_OI = FP_OI';end
                % for both VVA- VVA+
                if all(FP_OI~=1) % 沒有任何位置是 Pass, 只有 TIR 或是 Fail
                    limitedVA(:,3:4,whichWD) = -777;
                    break
                end
                while 1 % for VVA- [limitedVA(:,3,:)]
                    if FP_OI(1) == 1 % all 1: 777
                        limitedVA(:,3,whichWD) = 777;
                        break
                    end
                    indTIR = min(strfind(FP_OI',[-1 1])); % TIR 後 Pass: 紀錄
                    if ~isempty(indTIR)
                        limitedVA(1,3,whichWD) = VVA_list(indTIR);    % fail
                        limitedVA(2,3,whichWD) = VVA_list(indTIR+1);  % pass  
                        break
                    end
                    ind1 = min(strfind(FP_OI',[0,1])); % 正常情形 0 --> 1
                    limitedVA(1,3,whichWD) = VVA_list(ind1);    % fail
                    limitedVA(2,3,whichWD) = VVA_list(ind1+1);  % pass                   
                    break
                end
                while 1 % for VVA+ [limitedVA(:,4,:)]
                    if FP_OI(end) == 1
                        limitedVA(:,4,whichWD) = 777;
                        break
                    end
                    indTIR = max(strfind(FP_OI',[1 -1])); % TIR 後 Pass: 紀錄
                    if ~isempty(indTIR)
                        limitedVA(1,4,whichWD) = VVA_list(indTIR+1);    % fail
                        limitedVA(2,4,whichWD) = VVA_list(indTIR);      % pass  
                        break
                    end
                    ind2 = max(strfind(FP_OI',[1,0]));
                    limitedVA(1,4,whichWD) = VVA_list(ind2+1);    % pass
                    limitedVA(2,4,whichWD) = VVA_list(ind2);  % fail
                break
                end               
            end
            break
        end
        failNumberMatrix_WD(:,:,whichWD) = failNumberMatrix;
        failPassMatrix_WD(:,:,whichWD) = failPassMatrix;
        cprintf("完成 \n")
    end
    cprintf("\n");
    disp("Analysis Process Done!")
    %% write excel
    if writeExcel == 1
        cprintf("Writing Excel......");
        for whichWD2 = 1:length(WD_list)
            WD = WD_list(whichWD2);
            % 建立模板
            strtemp = ["HVA-";"HVA+";"VVA-";"VVA+"];   
            writematrix(strcat("(-- WD ",num2str(WD)," --)"),filename_excel,'Sheet',excelSheetName,'Range',strcat('A',num2str(1+10*(whichWD2-1))));
            writematrix(strtemp,filename_excel,'Sheet',excelSheetName,'Range',strcat('A',num2str(3+10*(whichWD2-1))));
            % 放入 Data
            writematrix(limitedVA(:,:,whichWD2)',filename_excel,'Sheet',excelSheetName,'Range',strcat('B',num2str(3+10*(whichWD2-1))));
            % 紀錄 fail RP
            writematrix("VVA/HVA",filename_excel,"Sheet",strcat("Raw(VD",num2str(WD),")"),"Range","A1")
            writematrix(HVA_list,filename_excel,"Sheet",strcat("Raw(VD",num2str(WD),")"),"Range","B1")
            writematrix(VVA_list',filename_excel,"Sheet",strcat("Raw(VD",num2str(WD),")"),"Range","A2")
            writematrix(failNumberMatrix_WD(:,:,whichWD2),filename_excel,"Sheet",strcat("Raw(VD",num2str(WD),")"),"Range","B2")
        end
        % 自動調整 寬
        cprintf("Excel AutoTune...");
        hExcel = actxserver('Excel.Application');
        hWorkbook = hExcel.Workbooks.Open(strcat(cd,"\",filename_excel)); 
        sheetNum = hWorkbook.Sheets.Count;
        for ex = 1:sheetNum
            FS_sheet = hWorkbook.Sheets.Item(ex);
            % alignment
            FS_sheet.Cells.EntireColumn.AutoFit;
            FS_sheet.Cells.EntireRow.AutoFit;
            FS_sheet.Cells.HorizontalAlignment = 3;
            FS_sheet.Cells.VerticalAlignment = 2;
            % select
            FS_sheet.Select;
            FS_sheet.Cells.Font.Name = 'Calibri';
            FS_sheet.Range("A1").Select;
        end
        hWorkbook.Save
        hWorkbook.Close
        hExcel.Quit
        cprintf("Done.\n");
    end
end
function failPixelNumber = getFailNumber(II)
    % 檢查 Viewing Zone
    if size(II,3) == 3
        redChannel = II(:, :, 1);
        greenChannel = II(:, :, 2);
        blueChannel = II(:, :, 3);
    elseif size(II,3) == 1
        redChannel = nan;
        greenChannel = nan;
        blueChannel = II;
    end
    
    mask = (redChannel >= 1 & greenChannel >= 1)... % 左右眼模式: 出現黃色 --> RP相撞
        | (blueChannel >= 2); % 單眼模式: 出現 255藍 --> RP相撞
    [rows, ~] = find(mask);
    failPixelNumber = length(rows);
end

function imageString = getImageStr(VD,VVA,HVA,filename,analyze_soft)
    switch analyze_soft
        case 0 % hardware
            VAPattern = caseInsensitivePattern(strcat("_WDR",num2str(VD),"_VVA",num2str(VVA),"_HVA",num2str(HVA),"_"));
        case 1 % software
            VAPattern = caseInsensitivePattern(strcat("_VD=",num2str(VD,'%07.2f'),"_VVA=",num2str(VVA,'%05.2f'),"_HVA=",num2str(HVA,'%+06.2f'),"_"));
    end
    VAContain = contains(filename,VAPattern);
    try
        imageString = filename{VAContain};
    catch
        beep;
        disp(VAPattern)
        switch analyze_soft
            case 0
                error("cannot extract VA value. (確保命名包含:... _WDR700_VVA30_HVA10_ ...)");
            case 1
                error("cannot extract VA value. (確保命名包含:... _VD=0400.00_VVA=30.00_HVA=+00.00_ ...)");
        end
    end
end