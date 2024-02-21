classdef Image < handle & matlab.mixin.Copyable
    %UNTITLED Summary of this class goes here
    %   Detailed explanation goes here
    
    properties
        name
        data
        shape    % 三維
    end
    
    methods
        function obj = Image(options)
            
            arguments
                options.name
                options.name_only = 0
                options.data
            end
            
            %% 兩者不能同時輸入
            if isfield(options,"name") && isfield(options,"data")
                error("檔名(name) 與 資料(data) 不能同時輸入");
            end

            %% 若輸入為 name --> 嘗試 imread
            if isfield(options,"name")
                obj.name = options.name;
                if options.name_only == 0
                    try
                        obj.data = imread(obj.name);
                    catch
                        error("找不到名為: '" + obj.name + " '的影像")
                    end
                elseif options.name_only == 1
                else
                    error("wrong input: name_only (can only be 0 or 1).")
                end
            end

            %% 若輸入為 data --> name 為空
            if isfield(options,"data")
                obj.data = options.data;
                obj.name = "";
            end

            %% size
            obj.update_size;
        end
        
        %% 
        function set_data(obj,value)
            obj.data = value;
            obj.update_size;
        end
        %% Other Function
        function status = isTIR_by_name(obj,str)
            % status: 0 (noTIR), 1 (TIR)
            containTIR = contains(obj.name,str);
            switch containTIR
                case 1
                    cprintf('key',"[info]: B 發生 TIR (continue to next)\n")
                    status = 1;
                case 0
                    status = 0;
                    % pass
                otherwise
                    cprintf('err',"[錯誤]: 無法偵測 B 的 TIR 項目\n")
                    cprintf('err',"系統停止\n")
                    throw(MException("MATLAB:test",""))
            end 
        end

        function status = equalto(obj, image)
            % status: 0 (different), 1 (same)
            if isa(image,'Image')
                status = isequal(obj.data,image.data);
            else
                status = isequal(obj.data,image);
            end
        end
        
        function update_size(obj)
            temp1 = [1,1,1];
            temp2 = size(obj.data);
            temp1(1:length(temp2)) = temp2;
            obj.shape = temp1;
        end

    end
end

