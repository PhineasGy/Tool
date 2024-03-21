classdef View < handle & matlab.mixin.Heterogeneous
    %UNTITLED Summary of this class goes here
    %   Detailed explanation goes here
    
    properties
        number
        resolution
        image
        arr_image
    end
    
    methods
        function obj = View(input)
            arguments
                input.number
                input.resolution
            end
            obj.number = input.number;
            obj.resolution = input.resolution;  % [ver,hor]
            obj.image = cell(1,obj.number);
        end
        function show(obj,ind)
            if isnumeric(ind)
                if 1 <= ind && ind <= obj.number
                    figure
                    imshow(obj.image{ind})
                else
                    error("找不到 View"+ind+" 的影像")
                end
            elseif isequal(ind,"arr")
                if ~isempty(obj.arr_image)
                    figure
                    imshow(obj.arr_image)
                else
                    error("無法顯示 arrange_image (arrange_image 為空)")
                end
            else
                error("error using View.show")
            end
        end
        function write(obj,varargin)
            if nargin == 1 % 寫出全部 View 影像
                for ii = 1:obj.number
                    imwrite(obj.image{ii},"View"+ii+".png");
                end
            elseif  nargin == 2 || nargin == 3  % 寫出指定 (第三項為檔名)
                if nargin == 3
                    str = varargin{2};
                else
                    str = "";
                end
                if isnumeric(varargin{1})
                    if 1 <= varargin{1} && varargin{1} <= obj.number
                        imwrite(obj.image{varargin{1}},"View"+varargin{1}+"_"+str+".png");
                    else
                        error("找不到 View"+ind+" 的影像")
                    end
                elseif isequal(varargin{1},"arr")
                    if ~isempty(obj.arr_image)
                        imwrite(obj.arr_image,"Arrange_"+str+".png");
                    else
                        error("無法寫出 arrange_image (arrange_image 為空)")
                    end
                else
                    error("error using View.write")
                end
            end
        end
        function createII(obj,options)
            arguments
                obj
                options.method
                options.only   % array. ex: [1,3]
            end
            only = 1:obj.number;
            if isfield(options,"only");only = options.only;end

            % current method: for LRA0
            switch options.method
                case "pixel"
                    obj.method_pixel(only);
                case "subpixel"
                    obj.method_subpixel(only);
            end

        end
        function method_pixel(obj,only)
            reverse = obj.number:-1:1;  % 第一個位置放最後一個 View，依此類推
            obj.arr_image = uint8(zeros([obj.resolution 3]));
            temp = size(obj.image{1},2);
            for ii = 1:temp             % view image size
                for vv = only           % (default) 1 : view number
                    ind = reverse(vv) + (obj.number) * (ii-1);
                    obj.arr_image(:,ind,:) = obj.image{vv}(:,ii,:);
                end
            end
        end
        function method_subpixel(obj,only)
            
            %% get index array
            array = nan(obj.number,3); % ex: V1 to V4, rgb
            % 起始位置: V4-R
            % 排列: RGB...
            ind = [obj.number,1];    % 不斷往右往上
            for ii = 1:obj.number*3
                array(ind(1),ind(2)) = ii;
                % 移動
                ind(1) = mod(ind(1) - 1,obj.number); if ind(1) == 0; ind(1) = obj.number;end
                ind(2) = mod(ind(2) + 1,3); if ind(2) == 0; ind(2) = 3;end
            end

            %% assignment
            % arrange image = (V) x (H) x 3 (V,H: 原圖解析度)
            reverse = obj.number:-1:1;  % 第一個位置放最後一個 View，依此類推
            rgb = 1:3;
            obj.arr_image = uint8(zeros([obj.resolution(1) obj.resolution(2) 3]));
            for ii = 1:obj.resolution(2)/(obj.number)           % view image size
                for vv = only                                   % (default) 1 : view number
                    interest_ind = array(vv,:);
                    for rgb = 1:3
                        try
                            if interest_ind(rgb) > obj.resolution(2); continue;end
                            obj.arr_image(:,interest_ind(rgb),rgb) = obj.image{vv}(:,ii,rgb);
                        catch
                            disp(":")
                        end
                    end
                end
                array = array + (obj.number);
            end
        end
    end
end

