classdef Mask
    %CRITICAL Summary of this class goes here
    %   Detailed explanation goes here
    
    properties
        panel_pixel_number_hor
        panel_pixel_number_ver
    end
    
    methods
        function obj = Mask(namedArgs)
            arguments
                namedArgs.panel_pixel_number_hor
                namedArgs.panel_pixel_number_ver
            end

            obj.panel_pixel_number_hor = namedArgs.panel_pixel_number_hor;
            obj.panel_pixel_number_ver = namedArgs.panel_pixel_number_ver;
        end
        
        function outputArg = method1(obj,inputArg)
            %METHOD1 Summary of this method goes here
            %   Detailed explanation goes here
            outputArg = obj.Property1 + inputArg;
        end
    end
end

