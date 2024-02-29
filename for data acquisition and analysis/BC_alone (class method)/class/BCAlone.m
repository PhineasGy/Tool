classdef BCAlone < handle
    
    properties
        INPUT
        OPTION
        CRITICAL
        MASK
    end
    properties
        % note: index 1 -> B, index 2 -> C, index 3 -> C Mask
        all_png
        target_png  % dynamic change
    end
    properties
        result_cell
        result_table
    end

    
    methods
        function obj = BCAlone(namedArgs)   % main program
            arguments
                namedArgs.INPUT
                namedArgs.OPTION
                namedArgs.CRITICAL
                namedArgs.MASK
            end
            obj.INPUT = namedArgs.INPUT;
            obj.OPTION = namedArgs.OPTION;
            obj.CRITICAL = namedArgs.CRITICAL;
            obj.MASK = namedArgs.MASK;

            %% pre-processing
            % index: B --> C --> Mask
            switch obj.OPTION.mask_mode
                case 0
                    num = 2;
                case 1
                    num = 3;
            end
            obj.all_png = cell(1,num);
            obj.target_png = cell(1,num);
            
            WD_list = obj.INPUT.WD_list;
            VVA_center = obj.INPUT.VVA_center;
            HVA_center = obj.INPUT.HVA_center;
            VVA_list = obj.INPUT.VVA_list;
            HVA_list = obj.INPUT.HVA_list;
            PS_list = obj.INPUT.PS_list;
            VA_mode = obj.OPTION.VA_mode;
            panel_pixel_number_ver = obj.MASK.panel_pixel_number_ver;
            panel_pixel_number_hor = obj.MASK.panel_pixel_number_hor;

            WD_num = length(WD_list);
            result_cell = cell(WD_num,length(VA_mode));
            %% processing
            
            %% step1: 選擇資料夾 B --> C--> Mask(optional)
            obj.select_directory;

            %% step2: VA loop
            for which_WD = 1:WD_num     % WD loop
                WD_now = WD_list(which_WD);
                PS_now = PS_list(which_WD);
                for which_VA_Term = VA_mode  % VA term loop
                    switch which_VA_Term
                        case 1      % HVA
                            VVA_list_now = linspace(VVA_center,VVA_center,length(HVA_list)); 
                            HVA_list_now = HVA_list;   
                            HVAColumn = HVA_list_now';
                        case 2      % VVA
                            VVA_list_now = VVA_list;                  
                            HVA_list_now = linspace(HVA_center,HVA_center,length(VVA_list));
                            VVAColumn = VVA_list_now';
                    end
                    VA_count = 0;
                    for which_VA = 1:length(VVA_list_now)  % VA loop
                        VA_count = VA_count + 1;
                        VVA_now = VVA_list_now(VA_count);
                        HVA_now = HVA_list_now(VA_count);
                        
                        %% Image Extraction
                        % VA pattern
                        switch obj.OPTION.II_mode
                            case "soft"
                               % EX:
                               % [BLP]_IPD=60.00_VD=0400.00_VVA=10.00_HVA=+00.00_PR=41.00_TIR=True.png
                                VA_pattern = strcat("_VD=",num2str(WD_now,"%07.2f"),...
                                                    "_VVA=",num2str(VVA_now,"%05.2f"),...
                                                    "_HVA=",num2str(HVA_now,"%+06.2f"));
                                % update: check PS for C
                                PR_pattern = strcat("_PR=",num2str(PS_now,"%05.2f"),"_");
                            case "hard"
                                % ignore
                        end
                        cprintf([1,0.5,0],strcat("[info]: 目標影像 ",VA_pattern,"\n"))
                        
                        % get target png/folder
                        for which_image = 1:num % order: B --> C --> Mask
                            all_png_name = string({obj.all_png{which_image}.name});
                            all_png_folder = string({obj.all_png{which_image}.folder});
                            switch which_image
                                case 1
                                    ind_temp1 = contains(all_png_name,VA_pattern);
                                otherwise   % C 和 Mask 均檢查 PS
                                    ind_temp1 = contains(all_png_name,VA_pattern) &...
                                        contains(all_png_name,PR_pattern);
                            end
                            target_png_name = all_png_name(ind_temp1);
                            target_png_folder = all_png_folder(ind_temp1);
                            if isempty(target_png_name)
                                switch which_image
                                    case 1
                                        errstr = "<BLP>" + VA_pattern;
                                    case 2
                                        errstr = "<C>" + VA_pattern + " " + PR_pattern;
                                    case 3
                                        errstr = "<Mask>" + VA_pattern + " " + PR_pattern;
                                end
                                error("[error] cannot find file with following pattern: " + errstr + "(系統停止)")
                            end
                            if length(target_png_name) ~= 1
                                target_png_name = target_png_name(1);  % 遇到兩個以上符合的檔名: 取第一個
                                target_png_folder = target_png_folder(1);
                            end
                            obj.target_png{which_image}.name = target_png_name;
                            obj.target_png{which_image}.folder = target_png_folder;
                        end
                        
                        %% check TIR (check B only)
                        % note: C 可能因為 B TIR 而無 data
                        if obj.OPTION.is_checking_TIR == 1
                            target_B_name = obj.target_png{1}.name;
                            isTIR = Image(name=target_B_name,name_only=1).isTIR_by_name("TIR=True");
                            switch isTIR
                                case 0  % 沒有 TIR
                                case 1  % 有 TIR
                                    cprintf('key',"[info]: B 發生 TIR (continue to next) [C 記為 no_info, B+C 記為 TIR]\n")
                                    result_cell{which_WD,which_VA_Term}.C_fail_list(VA_count) = "no_info";
                                    result_cell{which_WD,which_VA_Term}.BandC_fail_list(VA_count) = "TIR";
                                    continue
                            end
                        end

                        
                        %% Image Object
                        B_ImageObject = Image(name=fullfile(obj.target_png{1}.folder,obj.target_png{1}.name));
                        C_ImageObject = Image(name=fullfile(obj.target_png{2}.folder,obj.target_png{2}.name));
                        B = double(B_ImageObject.data);   % 理論上為紅綠圖
                        C = double(C_ImageObject.data);   % 理論上為紅綠圖
                        
                        %% C mask mode: 套上指定大小的 Mask
                        if obj.OPTION.mask_mode == 1
                            Mask_ImageObject = Image(name=fullfile(obj.target_png{3}.folder,obj.target_png{3}.name));
                            C_mask = double(Mask_ImageObject.data); % with content
                        
                            %% 非零邊界
                            C_mask_temp = C_mask(:,:,1) ~= 0 |C_mask(:,:,2) ~= 0|C_mask(:,:,3) ~= 0;
                            C_mask_temp2 = find(C_mask_temp==1);
                            [DD_row,DD_col] = ind2sub([panel_pixel_number_ver,panel_pixel_number_hor],...
                                C_mask_temp2);
                            row_min = min(DD_row);row_max = max(DD_row);
                            col_min = min(DD_col);col_max = max(DD_col);
                            C_mask = poly2mask([col_min,col_max,col_max,col_min],...
                                             [row_min,row_min,row_max,row_max],...
                                             panel_pixel_number_ver,panel_pixel_number_hor);
                            B = B.*C_mask;
                            C = C.*C_mask;
                            
                            % update image data (optional)
                            B_ImageObject.set_data(B);
                            C_ImageObject.set_data(C);
                            Mask_ImageObject.set_data(C_mask);
                        end
                        
                        %% image RGB value handle
                        maxRGB_B = max(B,[],"all");
                        B = B * 0.5;
                        maxRGB_C = max(C,[],"all");
                        C = C * 0.5;
                        if maxRGB_B ~= maxRGB_C
                            error(strcat("B and C RGB 最大值不相同: B:", num2str(maxRGB_B)," C:",num2str(maxRGB_C)));
                        end
                        
                        %% C Fail check
                        [C_fail_array,~] = find(C(:,:,1) >= 0.5 * maxRGB_C & C(:,:,2) >= 0.5 * maxRGB_C);
                        C_fail_count = length(C_fail_array);
                        disp(strcat("[info]: C fail number: ",num2str(C_fail_count)));
                        if obj.OPTION.is_checking_critical == 1                          
                            if obj.CRITICAL.critical_C_fail < C_fail_count
                                cprintf('key',"[result]: C Fail (黃過多) (continue to next) [C, B+C 記為 F1]\n")
                                result_cell{which_WD,which_VA_Term}.C_fail_list(VA_count) = "F1";
                                result_cell{which_WD,which_VA_Term}.BandC_fail_list(VA_count) = "F1";
                                continue
                            end
                        end
                        
                        %% B 填滿 check
                        [B_black_array,~] = find(B(:,:,1) == 0 & B(:,:,2) == 0 & B(:,:,3) == 0);
                        B_black_count = length(B_black_array);
                        disp(strcat("[info]: B black number: ",num2str(B_black_count)));
                        if obj.OPTION.is_checking_critical == 1
                            if obj.CRITICAL.critical_B_black < B_black_count
                                cprintf('key',"[result]: B Fail (黑過多) (continue to next) [B+C 記為 F2]\n")
                                result_cell{which_WD,which_VA_Term}.C_fail_list(VA_count) = string(C_fail_count);
                                result_cell{which_WD,which_VA_Term}.BandC_fail_list(VA_count) = "F2";
                                continue
                            end
                        end
                        
                        %% B + C Fail check
                        B_C = B + C;
                        [B_and_C_fail_array,~] = find(B_C(:,:,1) >= 0.5 * maxRGB_C & B_C(:,:,2) >= 0.5 * maxRGB_C);
                        B_and_C_fail_count = length(B_and_C_fail_array);
                        disp(strcat("[info]: B + C fail number: ",num2str(B_and_C_fail_count)));
                        if obj.OPTION.is_checking_critical == 1
                            if obj.CRITICAL.critical_BC_fail < B_and_C_fail_count
                                cprintf('key',"[result]: B + C Fail (黃過多) (continue to next) [B+C 記為 F3]\n")
                                result_cell{which_WD,which_VA_Term}.C_fail_list(VA_count) = string(C_fail_count);
                                result_cell{which_WD,which_VA_Term}.BandC_fail_list(VA_count) = "F3";
                                continue
                            end
                        end
                        
                        %% Pass: 
                        if obj.OPTION.is_checking_critical == 1
                            cprintf('key',"[result]: Pass\n")
                        end
                        result_cell{which_WD,which_VA_Term}.C_fail_list(VA_count) = string(C_fail_count);
                        result_cell{which_WD,which_VA_Term}.BandC_fail_list(VA_count) = string(B_and_C_fail_count);
                    end % VA Loop
                end
            end
            obj.result_cell = result_cell;
            cprintf("=== 分析完成 ===\n")

            %% step 3: create table
            if obj.OPTION.is_creating_table == 1
                finalTable(length(WD_list)).BC = nan;
                
                tableBase = cell(length(VVAColumn)+1,length(HVAColumn)+1);
                tableBase(:) = {""};
                tableBase(1,1) = {"VVA/HVA"};
                tableBase(2:end,1) = num2cell(VVAColumn);
                tableBase(1,2:end) = num2cell(HVAColumn);
                
                for whichWD = 1:length(WD_list)
                    tableBC = tableBase;
                    tableBC(1+find(VVAColumn == VVA_center,1,"first"),2:end) = num2cell(result_cell{whichWD,1}.C_fail_list);
                    tableBC(1+find(VVAColumn == VVA_center,1,"last"),2:end) = num2cell(result_cell{whichWD,1}.BandC_fail_list);
                    tableBC(2:end,1+find(HVAColumn == HVA_center,1,"first")) = num2cell(result_cell{whichWD,2}.C_fail_list);
                    tableBC(2:end,1+find(HVAColumn == HVA_center,1,"last")) = num2cell(result_cell{whichWD,2}.BandC_fail_list);
                    finalTable(whichWD).BC = tableBC;
                end
                
                obj.result_table = finalTable;
            end
        end
        
        function select_directory(obj)
            %% preprocessing
            str_1 = ["選擇 B 資料夾","選擇 C 資料夾","選擇 Mask 資料夾"];
            num = length(obj.all_png);

            %% processing
            for which_image = 1 : num
                currentWorkspace = cd;
                dir_str = uigetdir(cd,str_1(which_image));
                cd(dir_str)
                obj.all_png{which_image} = dir("**/*.png");    % 選取包含子目錄的 png 檔案
                cd(currentWorkspace)
            end
        end

        function open_table(obj)    %#ok<MANU> % for variable "BC_ALONE" only
            openvar BC_ALONE.result_table
        end
    end
end

