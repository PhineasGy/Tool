classdef Option
    %UNTITLED Summary of this class goes here
    %   Detailed explanation goes here
    
    properties
        VA_mode         % "1, 2, 1:2" [HVA, VVA, HVA --> VVA]
        mask_mode       % regional VZA
        II_mode         % "hard" "soft"
        is_creating_table
        is_checking_TIR
        is_checking_critical
    end
    
    methods
        function obj = Option(namedArgs)
            arguments
                namedArgs.VA_mode
                namedArgs.mask_mode
                namedArgs.II_mode
                namedArgs.is_creating_table
                namedArgs.is_checking_TIR
                namedArgs.is_checking_critical
            end
            
            obj.VA_mode = namedArgs.VA_mode;
            obj.mask_mode = namedArgs.mask_mode;
            obj.II_mode = namedArgs.II_mode;
            obj.is_creating_table = namedArgs.is_creating_table;
            obj.is_checking_TIR = namedArgs.is_checking_TIR;
            obj.is_checking_critical = namedArgs.is_checking_critical;
        end
        
        function outputArg = method1(obj,inputArg)
            %METHOD1 Summary of this method goes here
            %   Detailed explanation goes here
            outputArg = obj.Property1 + inputArg;
        end
    end
end

