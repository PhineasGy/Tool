%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% version: 20220923 (chinese comment support)
close all; 
clc;clear; 
%%

frameCount = 4;

% comment
addComment = 1;
comment = strings(1,frameCount);
comment(1) = "LLDefault OE";
comment(2) = "LLShrink OE";
comment(3) = "LLDefault UE";
comment(4) = "LLShrink UE";
fontSize = 200; % maximum:200

boxColor ='yellow';
boxOpacity = 1; % (0-1) no box if set to 0
textColor = 'black';

% gif
delayTime = 1.5; % seconds
filename = "OOVA_II_RDP40.gif";

%% 
FrameArray = cell(1,frameCount);

% Frame1 = im2uint8(imread("AAMask1_40_PS0_LLDefault_OE.png"));
% Frame2 = im2uint8(imread("AAMask1_40_PS0_LLShrink_UE.png"));

% 選擇圖檔
imagePathname=[];
for chooseImage = 1:frameCount

    if addComment == 1
        titleName = strcat("原圖: (",comment(chooseImage),")");
    elseif addComment == 0
        titleName = "原圖";
    end
    [imageFilename, imagePathname] = uigetfile({strcat(imagePathname,'*.png;','*.bmp')}, titleName);
    if ~ischar(imagePathname) 
        return;end
    imageFilepath = fullfile(imagePathname, imageFilename);
    FrameArray{chooseImage} = im2uint8(imread(imageFilepath));
end

%% 選擇 comment 位置
if addComment == 1
    I = FrameArray{1};
    pass = 0;
    while pass == 0
        F1 = figure;
        imshow(I)
        title("選擇 comment 的左上位置")
        roi = drawpoint;
        try
            position = [roi.Position(1),roi.Position(2)];
        catch
            disp("System Stopped.")
            return
        end
        close(F1)
        IAfterLabel = insertText(I,position,comment(1),'FontSize',fontSize,'Font','SimSun',...
            'BoxColor',boxColor,'BoxOpacity',boxOpacity,'TextColor',textColor);
        F2 = figure;
        imshow(IAfterLabel)       
        anwser = questdlg('Do it again? (若要調整字體大小請取消重來)', ...
	                            'Quick Ask', ...
	                            '再決定一次位置','繼續下一步','再決定一次位置');
        if isempty(anwser)
            disp("System Stopped.")
            close(F2)
            return
        elseif anwser == "再決定一次位置"
        elseif anwser == "繼續下一步"
            pass = 1;
        end
        close(F2)
    end
    % add comment to images
    for whichFrame = 1:frameCount
        I = FrameArray{whichFrame};
        FrameArray{whichFrame} = insertText(I,position,comment(whichFrame),'FontSize',fontSize,'Font','SimSun',...
            'BoxColor',boxColor,'BoxOpacity',boxOpacity,'TextColor',textColor);figure;
        imshow(FrameArray{whichFrame})
    end

end

%%

% RGB2IND
% matlab gif : index image support only.
% FrameData Creation
for whichFrame = 1:frameCount

    if size(FrameArray{whichFrame},3) == 1
        FrameArray{whichFrame} = repmat(FrameArray{whichFrame},[1,1,3]);
    end
    [A,map] = rgb2ind(FrameArray{whichFrame},256);
    if whichFrame == 1
        imwrite(A,map,filename,'gif','LoopCount',Inf,'DelayTime',delayTime);
    else
        imwrite(A,map,filename,'gif','WriteMode','append','DelayTime',delayTime);
    end
end
% 
disp("Gif 完成")
