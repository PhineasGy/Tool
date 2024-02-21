function CIS()
    % 選擇影像1
    Image_pathname=[];
    [Image_filename, Image_pathname] = uigetfile({'*.*'}, '選擇影像1');
    if ~ischar(Image_pathname) 
    error("未選擇影像");end
    Image_filepath = fullfile(Image_pathname, Image_filename);
    try
        Image1=im2double(imread(Image_filepath));
    catch
        error("讀取影像失敗");
    end
    % 選擇影像2
    [Image_filename, Image_pathname] = uigetfile({'*.*'}, '選擇影像2');
    if ~ischar(Image_pathname) 
    error("未選擇影像");end
    Image_filepath = fullfile(Image_pathname, Image_filename);
    try
        Image2=im2double(imread(Image_filepath));
    catch
        error("讀取影像失敗");
    end
    % 比較兩者
    if isequal(Image1,Image2)
        beep
        disp("兩張影像相同無異")
    else
        beep
        disp("兩張影像有差異")
    end
    
end