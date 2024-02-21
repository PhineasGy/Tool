function [imageCroppedArray,returnCheck] = imcropGY(image,sizeArray,offsetArray,customLine,varargin)
% Version: 20220617
% Usage:
% 1. imcropGY(): 手動選取圖片手動做裁切 (multiselectFig,imcrop,"_Cropped")
% 2. imcropGY(image,[sizeVer,sizeHor],[offserVer,offsetHor],customLine,Name,Value)
%   image (optional): 準備做裁切的影像. image == nan: 手動選取圖片.
%   [sizeVer,sizeHor] (optional): 裁切影像大小. Default:[size(image,1),size(image,2)]
%   [offserVer,offsetHor](optional): 裁切影像中心位移. Default:[0,0]
%   customLine (optional): 裁切後圖片檔名. Default:"_Cropped"
% 3. imcropGY(Name,Value)
%   Name-Value Arguments:
%       Show:'on','off'(Default)
%       Write:'on'(Default),'off'
%
% disp(nargin)
returnCheck = 0;
numMainInput = nargin - length(varargin);
% numNameValue = length(varargin);
%% input 參數處理
defaultImage = nan; % 手動選圖
defaultSizeArray = []; %
defaultOffsetArray = []; %
defaultCustomLine = "_Cropped";
defaultShow = "off";
defaultWrite = "on";
defaultManualCrop = "off";
expectedSwitch = {'on','off'};

switch numMainInput
    case 0
        image = nan;
        sizeArray = [];
        offsetArray = [];
        customLine = "_Cropped";
        defaultManualCrop = "on";
    case 1
        sizeArray = [];
        offsetArray = [];
        customLine = "_Cropped";
        defaultManualCrop = "on";
    case 2
        offsetArray = [];
        customLine = "_Cropped";
    case 3
        customLine = "_Cropped";
    case 4
    otherwise
        error("too many input. :<")
end

p = inputParser;

validImage = @(x) any(all(isnan(x),'all') | all((isnumeric(x) & (x>=0)),'all'),'all');
validSizeArray = @(x) (length(x)==2 && all(logical(~rem(x,1))) && all(x>0)) || isequal(x,[]);
validOffsetArray = @(x) (length(x)==2 && all(logical(~rem(x,1)))) || isequal(x,[]);
validCustomLine = @(x) ischar(x) || isstring(x);
validSwitch = @(x) any(validatestring(x,expectedSwitch));

addOptional(p,'image',defaultImage,validImage);
addOptional(p,'sizeArray',defaultSizeArray,validSizeArray);
addOptional(p,'offsetArray',defaultOffsetArray,validOffsetArray);
addOptional(p,'customLine',defaultCustomLine,validCustomLine);
% Name Value Pair
addParameter(p,'Show',defaultShow,validSwitch);
addParameter(p,'Write',defaultWrite,validSwitch);
addParameter(p,'ManualCrop',defaultManualCrop,validSwitch);

parse(p,image,sizeArray,offsetArray,customLine,varargin{:}) % 解析輸入

%% input parameter assignment %%
image = p.Results.image;
sizeArray = p.Results.sizeArray;
offsetArray = p.Results.offsetArray;
customLine = p.Results.customLine;
show = p.Results.Show;
write = p.Results.Write;
manualCrop = p.Results.ManualCrop;

% 判定選圖
if isnan(image) % 自選圖 (支援 png, bmp)
    imagePathname=[];
    [imageFilename, imagePathname] = uigetfile({strcat(imagePathname,'*.png;',imagePathname,'*.bmp')}, '欲做裁切的圖','MultiSelect', 'on');
    if ~ischar(imagePathname);returnCheck = 1;return;end
    if ischar(imageFilename)
        totalNumFile = 1;
    else
        totalNumFile = length(imageFilename);
    end
    imageCroppedArray = cell(totalNumFile,1);
    for filenum = 1:totalNumFile % 選多圖
        if totalNumFile==1
            name = imageFilename;
            imageFilepath = fullfile(imagePathname, name);
        else
            name = imageFilename{filenum};
            imageFilepath = fullfile(imagePathname, name);
        end
        image2Crop = imread(imageFilepath);
        
        if strcmpi(manualCrop,"on") % 手動切
            doItAgain = 1;
            if filenum == 1
                while (doItAgain)
                    f1 = figure('Name','Crop Your Image!');
                    [imageCropped,rect] = imcrop(image2Crop);
                    if isempty(imageCropped);disp("System Stopped.");returnCheck = 1;return;end
                    f2 = figure('Name','Cropped Image');
                    imshow(imageCropped)
                    anwser = questdlg('Is it okay?', ...
	                    'Quick Ask', ...
	                    'Meh~','Nah~','Meh~');
                    if isempty(anwser)
                        disp("System Stopped.")
                        returnCheck = 1;
                        return
                    elseif anwser == "Meh~"
                        doItAgain = 0;
                    end
                    try close(f1);catch;end
                    try close(f2);catch;end
                end
                try close(f1);catch;end
                try close(f2);catch;end
            else % 其他張沿用第一張切圖範圍
                imageCropped = imcrop(image2Crop,rect);
            end
        elseif strcmpi(manualCrop,"off") % 自動切
            [verSize,horSize,~] = size(image2Crop);
            if isequal(sizeArray,[])
                croppedVerSize = verSize;
                croppedHorSize = horSize;
            else
                croppedVerSize = sizeArray(1);
                croppedHorSize = sizeArray(2);
            end
            if isequal(offsetArray,[])
                centerOffsetVer = 0;
                centerOffsetHor = 0;
            else
                centerOffsetVer = offsetArray(1);
                centerOffsetHor = offsetArray(2);
            end    
            pixelLeftUp = [horSize*0.5+centerOffsetHor-croppedHorSize*0.5+1,verSize*0.5+centerOffsetVer-croppedVerSize*0.5+1];
            try
                imageCropped = imcrop(image2Crop,[pixelLeftUp(1),pixelLeftUp(2),croppedHorSize-1,croppedVerSize-1]);
            catch
                error("Error Cropping the image. Make sure the size and offset value is valid.")
            end
        end
        % 秀圖寫圖
        if strcmpi(show,"on")
            figure;
            imshow(imageCropped)
        end
        if strcmpi(write,"on")
            imwrite(imageCropped,strcat(erase(imageFilepath,[".png",".bmp"]),customLine,".png"));
        end
        imageCroppedArray{filenum} = imageCropped;
    end
else % 對 image 進行裁切
    imageCroppedArray = nan;
    image2Crop = image;
    if strcmpi(manualCrop,"on") % 手動切
    doItAgain = 1;
        while (doItAgain)
            f1 = figure('Name','Crop Your Image!');
            [imageCropped,~] = imcrop(image2Crop);
            if isempty(imageCropped);disp("System Stopped.");returnCheck = 1;return;end
            f2 = figure('Name','Cropped Image');
            imshow(imageCropped)
            anwser = questdlg('Is it okay?', ...
                'Quick Ask', ...
                'Meh~','Nah~','Meh~');
            if isempty(anwser)
                disp("System Stopped.")
                returnCheck = 1;
                return
            elseif anwser == "Meh~"
                doItAgain = 0;
            end
            try close(f1);catch;end
            try close(f2);catch;end
        end
        try close(f1);catch;end
        try close(f2);catch;end
    elseif strcmpi(manualCrop,"off") % 自動切
        [verSize,horSize,~] = size(image2Crop);
        if isequal(sizeArray,[])
            croppedVerSize = verSize;
            croppedHorSize = horSize;
        else
            croppedVerSize = sizeArray(1);
            croppedHorSize = sizeArray(2);
        end
        if isequal(offsetArray,[])
            centerOffsetVer = 0;
            centerOffsetHor = 0;
        else
            centerOffsetVer = offsetArray(1);
            centerOffsetHor = offsetArray(2);
        end    
        pixelLeftUp = [horSize*0.5+centerOffsetHor-croppedHorSize*0.5+1,verSize*0.5+centerOffsetVer-croppedVerSize*0.5+1];
        try
            imageCropped = imcrop(image2Crop,[pixelLeftUp(1),pixelLeftUp(2),croppedHorSize-1,croppedVerSize-1]);
        catch
            error("Error Cropping the image. Make sure the size and offset value is valid.")
        end
    end
    % 秀圖寫圖
    if strcmpi(show,"on")
        figure;
        imshow(imageCropped)
    end
    if strcmpi(write,"on")
        imwrite(imageCropped,strcat("Image",customLine,".png"));
    end
    imageCroppedArray = imageCropped;
end
end