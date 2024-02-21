%% General
clear
clc;
%%
stringR = "hard_noSeg_(升降K)";
stringG = "soft_noSeg";

withoutYellow = 0;

%% Read Image
[file,path] = uigetfile('*.*',strcat("R: ",stringR));
Red = imread(fullfile(path,file));
[file,path] = uigetfile('*.*',strcat("G: ",stringG));
Green = imread(fullfile(path,file));

%% size check
sizeX_R = size(Red,1);
sizeY_R = size(Red,2);
sizeX_G = size(Green,1);
sizeY_G = size(Green,2);
if sizeX_R~= sizeX_G || sizeY_R~= sizeY_G  
    error("Size for Red and Green is different.")
end

%% combine
combine = uint8(zeros(sizeX_R,sizeY_R,3));
combine(:,:,1) = Red(:,:,1);
combine(:,:,2) = Green(:,:,1);

if withoutYellow == 1
    combine_without_yellow = combine;
    temp = (combine(:,:,1) == 255) & (combine(:,:,2) == 255);
    combine_without_yellow(repmat(temp,[1,1,3])) = 0;
    imwrite(combine_without_yellow,strcat("combine_",stringR,"(R)_",stringG,"(G)_without_yellow.png"));
elseif withoutYellow == 0
    imwrite(combine,strcat("combine_",stringR,"(R)_",stringG,"(G).png"));
end
