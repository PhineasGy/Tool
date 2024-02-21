classdef Input

    properties % public
        VVA_center
        HVA_center
        WD_list
        VVA_list
        HVA_list
        OPTION
    end
    
    methods
        function obj = Input(namedArgs)
    
            arguments
             namedArgs.VVA_center
             namedArgs.HVA_center
             namedArgs.WD_list
             namedArgs.VVA_list
             namedArgs.HVA_list
             namedArgs.OPTION
            end

            obj.VVA_center = namedArgs.VVA_center;
            obj.HVA_center = namedArgs.HVA_center;
            obj.WD_list = namedArgs.WD_list;
            obj.VVA_list = namedArgs.VVA_list;
            obj.HVA_list = namedArgs.HVA_list;
            obj.OPTION = namedArgs.OPTION;
        end

        
        function outputArg = method1(obj,inputArg)
            %METHOD1 Summary of this method goes here
            %   Detailed explanation goes here
            outputArg = obj.Property1 + inputArg;
        end
    end
end

