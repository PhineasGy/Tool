%% Combine Tool: 
% author: GY
% 兩圖比較 (報告用)
clear
clc
%% 使用者輸入

add_directly = 0;       % 是否兩圖直接相加
% if not:
image_custom = 0;
    % default: 第一張影像轉為 red * 1
    % default: 第二張影像轉為 green * 1
    first_red = 1;
        first_red_factor = 1;
    first_green = 0;
        first_green_factor = 1;
    first_blue = 0;
        first_blue_factor = 1;
    second_red = 0;
        second_red_factor = 1;
    second_green = 1;
        second_green_factor = 1;
    second_blue = 0;
        second_blue_factor = 1;
% uint8: 255 + 255 --> 255; double: 1 + 1 --> 2 --> 1
combine_greater_than_1_equal_to_1 = 1;  % 主要針對 image_custom == 1 做處理

% 輸出檔名: 
% if add_directly == 0
% case 1: (conbine)_(first[1,0,0])_(second[0,1,0])_extra_string... --> if string == ""
% case 2: (conbine)_(first_string[1,0,0])_(second_string[0,1,0])_extra_string...--> if string ~= ""
% if add_directly == 1
% case 1: (combine)_(first)_(second)__extra_string --> if string == ""
% case 2: (combine)_(first_string)_(second_string)__extra_string --> if string ~= ""
first_string = "LTRP P1.002 A1.001";
second_string = "MLRP P1.001 A1.001";
extra_string = "L21 VD700 OOVA30 mideye PS15";
store_at = 0;   % 0: 目前路徑，1: 第一張圖路徑，2: 第二章圖路徑

%% select image
path_current = cd;
[file1,path1] = uigetfile("*.png",strcat("選擇第一張影像: ",first_string),"MultiSelect","off");
[file2,path2] = uigetfile(strcat(path1,"*.png"),strcat("選擇第二張影像: ",second_string),"MultiSelect","off");
first_image = imread(fullfile(path1,file1));
second_image = imread(fullfile(path2,file2));

% image handle: size check
[sizever_1,sizehor_1,size3D_1] = size(first_image);
[sizever_2,sizehor_2,size3D_2] = size(second_image);
if ~(sizever_1 == sizever_2 && sizehor_1 == sizehor_2)
    error("[error] 所選兩張影像大小不同，無法做合併 [系統停止]")
end
% image handle: to 24 bits
if size3D_1 == 1
    first_image = repmat(first_image,[1,1,3]);
end
if size3D_2 == 1
    second_image = repmat(second_image,[1,1,3]);
end

%% string handle
first_string = strcat("(first)",first_string);
second_string = strcat("(second)",second_string);

first_string;
second_string;
switch add_directly
    case 0
        switch image_custom
            case 0
                first_string = first_string.append("[1,0,0]");
                second_string = second_string.append("[0,1,0]");
            case 1
                first_string = first_string.append(strcat("[",nums2str(first_red*first_red_factor),","...
                    ,nums2str(first_green*first_green_factor),","...
                    ,nums2str(first_blue*first_blue_factor),"]"));
                second_string = second_string.append(strcat("[",nums2str(second_red*second_red_factor),","...
                    ,nums2str(second_green*second_green_factor),","...
                    ,nums2str(second_blue*second_blue_factor),"]"));
        end
    case 1
end
if extra_string == "" || isempty(extra_string)
    extra_string = [];
end
output_string = strjoin(["(combine)",first_string,second_string,extra_string],"_");
output_string = output_string.append(".png");

%% 影像處理
switch add_directly
    case 0  % 處理後相加
        combine_image = zeros(sizever_1,sizehor_1,3);
        first_image_handle = zeros(sizever_1,sizehor_1,3);
        secomd_image_handle = zeros(sizever_1,sizehor_1,3);
        % first image handle
        first_image_nonzeros = first_image(:,:,1)~=0 | first_image(:,:,2)~=0 | first_image(:,:,3)~=0;
        second_image_nonzeros = second_image(:,:,1)~=0 | second_image(:,:,2)~=0 | second_image(:,:,3)~=0;
        if image_custom == 1
            if first_red == 1
                first_red_image = first_image_nonzeros * first_red_factor;
            end
            if first_green == 1
                first_green_image = first_image_nonzeros * first_green_factor;
            end
            if first_blue == 1
                first_blue_image = first_image_nonzeros * first_blue_factor;
            end
            first_image_handle(:,:,1) = first_red_image;
            first_image_handle(:,:,2) = first_green_image;
            first_image_handle(:,:,3) = first_blue_image;
            
            if second_red == 1
                second_red_image = second_image_nonzeros * second_red_factor;
            end
            if second_green == 1
                second_green_image = second_image_nonzeros * second_green_factor;
            end
            if second_blue == 1
                second_blue_image = second_image_nonzeros * second_blue_factor;
            end
            second_image_handle(:,:,1) = second_red_image;
            second_image_handle(:,:,2) = second_green_image;
            second_image_handle(:,:,3) = second_blue_image;
            % [0-1] double
            combine_image = first_image_handle + second_image_handle;
            if combine_greater_than_1_equal_to_1 == 1
                combine_image(combine_image>1) = 1;
            end
        elseif image_custom == 0
            % [0-255] unit8
            combine_image(:,:,1) = first_image_nonzeros;
            combine_image(:,:,2) = second_image_nonzeros;
        end
    case 1  % 直接相加
        combine_image = first_image + second_image;
        % [0-255] unit8
end

%%  輸出
switch store_at
    case 0
        output_path = path_current;
    case 1
        output_path = path1;
    case 2
        output_path = path2;
end
imwrite(combine_image,fullfile(output_path,output_string))

cprintf('key',"程序完成\n")
beep