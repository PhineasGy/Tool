clear
clc

%% 使用者輸入
VVA_center = 30;
HVA_center = 0;

%% 工作表名稱
% 可自行修改，確保 Excel 工作表名稱有符合
excel_name = "0206 [Gy] 裕度 (L21A1-6).xlsx";
sheet_name_sort = "分析裕度整理";
sheet_name_VD = ["分析裕度 (VD400) (13)","分析裕度 (VD500) (13)","分析裕度 (VD600) (13)","分析裕度 (VD700) (13)"];

%% 裕度邊界, FailNumber 位置
% 可自行修改
% EX: D18 to G21 (4x4)
border_ind = ["D18","G21"];
% border_ind = ["L18","O21"];
% border_ind = ["T18","W21"];
% border_ind = ["AB18","AE21"];

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% fail numbaer table, VZA(critical median) table
failNumber_table = -1 * ones(4,4);
failNumber_table = num2cell(failNumber_table);
critical_median_table = -1 * ones(4,4);

%% VD_list
sheet_VD_list = cell(1,4);
center_row_list = nan(1,4);
center_col_list = nan(1,4);
VDCount = 0;
for whichVD = [400,500,600,700]
    VDCount = VDCount + 1;
    sheet_VD_list{VDCount} = readcell(excel_name,'Sheet',sheet_name_VD(VDCount));
    sheet_VD = sheet_VD_list{VDCount};
    % find center position (B+C)
    center_row = find(cellfun(@(x) isequal(x,VVA_center),sheet_VD(:,1)));
    center_row_list(VDCount) = center_row(2);
    center_column = find(cellfun(@(x) isequal(x,HVA_center),sheet_VD(1,:)));
    center_col_list(VDCount) = center_column(2);
end

%% fail number 
sheet_sort = table2array(readtable(excel_name,'Sheet',sheet_name_sort,'Range',border_ind(1)+":"+border_ind(2)));
VDCount = 0;
for whichVD = [400,500,600,700]
    VDCount = VDCount + 1;
    sheet_VD = sheet_VD_list{VDCount};
    center_row = center_row_list(VDCount);
    center_column = center_col_list(VDCount);

    % order: [VVA- VVA+ HVA- HVA+]
    for whichVATerm = 1:4
        border = sheet_sort(whichVATerm,VDCount);
        switch whichVATerm
            case {1,2} % VVA: 找 Column 1
                target_row = find(cellfun(@(x) isequal(x,border),sheet_VD(:,1)));
                if length(target_row) == 2; target_row = target_row(2);end    % center: 取後者 (B+C)
                if isempty(target_row)  % 分析裕度資料不足
                    failNumber_table{whichVATerm,VDCount} = 'OR';  % "Out of Range"
                    continue
                end
                temp = sheet_VD{target_row,center_column};
                failNumber_table{whichVATerm,VDCount} = temp;
            case {3,4} % HVA: 找 Row 1
                target_col = find(cellfun(@(x) isequal(x,border),sheet_VD(1,:)));
                if length(target_col) == 2; target_col = target_col(2);end    % center: 取後者 (B+C)
                if isempty(target_col)  % 分析裕度資料不足
                    failNumber_table{whichVATerm,VDCount} = 'OR';
                    continue
                end
                temp = sheet_VD{center_row,target_col};
                failNumber_table{whichVATerm,VDCount} = temp;
        end
    end
end

%% critical median
% rule: 第一個 fail number 必須  > median --> 否則為 TIR or OR
temp2 = {failNumber_table{:}}; %#ok<CCAT1> % 變為一維
temp2(cellfun(@(x) ~isnumeric(x),temp2))=[]; % 去除 "TIR" "OR"
temp2 = cell2mat(temp2);
critical_median = median(temp2);

VDCount = 0;
for whichVD = [400,500,600,700]
    VDCount = VDCount + 1;
    sheet_VD = sheet_VD_list{VDCount};
    center_row = center_row_list(VDCount);
    center_column = center_col_list(VDCount);

    % order: [VVA- VVA+ HVA- HVA+]
    for whichVATerm = 1:4
        switch whichVATerm
            case 1  % VVA-
                data = sheet_VD(2:end,center_column);
                VA_list = sheet_VD(2:end,1);
            case 2  % VVA+
                data = sheet_VD(end:-1:2,center_column);
                VA_list = sheet_VD(end:-1:2,1);
            case 3  % HVA-
                data = sheet_VD(center_row,2:end);
                VA_list = sheet_VD(1,2:end);
            case 4  % HVA+
                data = sheet_VD(center_row,end:-1:2);
                VA_list = sheet_VD(1,end:-1:2);
        end

        %% 不可有 missing
        missing_check = any(cellfun(@any, (cellfun(@(x) ismissing(x),data,'UniformOutput',false))));
        if missing_check; error("[error]: data contain 'missing'");end

        %% 建立 0 1 陣列
        % "NaN": TIR --> 1
        % data > median --> 1
        % data < median --> 0
        data_check = data;
        data_check(cellfun(@(x) isequal(x,'TIR'),data_check)) = {critical_median+1};    % 使其可以在下下行變為 1
        data_check(cellfun(@(x) le(x,critical_median),data_check)) = {0};           % lower than and equal (le)
        data_check(cellfun(@(x) gt(x,critical_median),data_check)) = {1};           % greater than (gt)
        data_check = cell2mat(data_check);

        % 找第一個 1 --> 0 的位置
        if size(data_check,1)~=1; data_check = data_check';end
        target_ind = strfind(data_check,[1 0]);
        if isempty(target_ind)  % 全為 0 or 1 (通常可能是全為 0)
            % 先取最邊界 (待討論)
            critical_median_table(whichVATerm,VDCount) = VA_list{1};
        else
            target_ind = target_ind(1) + 1; % 找 0 的位置
            critical_median_table(whichVATerm,VDCount) = VA_list{target_ind};
        end
    end
end

%%
open failNumber_table
open critical_median_table