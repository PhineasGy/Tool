classdef ALM_Table < handle
    properties
        tables = cell(0)
    end
    
    methods
        function obj = ALM_Table()
        end
        function add(obj,table)
            obj.tables{end+1} = table;
        end
        function target = find(obj,index,options)
            arguments
                obj
                index = 1 % 第幾個 table (預設 1)
                options.col
                options.row
            end
            target_table = obj.tables{index};
            target_table_data = target_table.Variables;
            col_name = target_table.Properties.VariableNames;
            row_name = target_table.V_L;
            if isfield(options,"col")
                col_chcek = cellfun(@(x) isequal(x,string(options.col)),col_name);
            else
                col_chcek = true(1,length(col_name));
            end
            if isfield(options,"row")
                row_chcek = cellfun(@(x) isequal(x,string(options.row)),row_name);
            else
                row_chcek = true(length(row_name),1);
            end
            target = target_table_data(row_chcek,col_chcek);
            if isempty(target);disp("cannot find target.");end
        end
    end
end

