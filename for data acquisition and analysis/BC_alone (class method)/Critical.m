classdef Critical
    %CRITICAL Summary of this class goes here
    %   Detailed explanation goes here
    
    properties
        C_fail
        B_black
        BC_fail
    end
    
    methods
        function obj = Critical(namedArgs)
            arguments
                namedArgs.critical_C_fail
                namedArgs.critical_B_black
                namedArgs.critical_BC_fail
            end

            obj.C_fail = namedArgs.critical_C_fail;
            obj.B_black = namedArgs.critical_B_black;
            obj.BC_fail = namedArgs.critical_BC_fail;
        end
        
        function outputArg = method1(obj,inputArg)
            %METHOD1 Summary of this method goes here
            %   Detailed explanation goes here
            outputArg = obj.Property1 + inputArg;
        end
    end
end

