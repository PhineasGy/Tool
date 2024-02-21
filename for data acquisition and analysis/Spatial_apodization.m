%% Spatial apodization Matlab version (空間切趾 For LT Sim)　Since 20210408
% version: 2023011701
% content: back to 13.3 Portrait setup
%%
close all; clc; clear;
%% 使用者輸入
ZerotoOne = 1;                   % 是否將 Matrix 中的 0 轉為 1
HorSize = 165.24;                % (mm)
VerSize = 293.76;                % (mm)

%% Processing %%
%% 讀圖
Image_pathname=[];
[Image_filename, Image_pathname] = uigetfile({strcat(Image_pathname,'*.png;',Image_pathname,'*.bmp')}, '欲做空間切趾的II圖','MultiSelect', 'on');
if ~ischar(Image_pathname) 
    return;end
if ischar(Image_filename)
    totalnum_file=1;
else
    totalnum_file=length(Image_filename);
end
%%
tic
for filenum=1:totalnum_file
if totalnum_file==1
    name=Image_filename;
    Image_filepath = fullfile(Image_pathname, Image_filename);
    OutputName=strcat(erase(Image_filename,[".png",".bmp"]),".txt");
    OutputName=fullfile(Image_pathname, OutputName);
    fileID = fopen(OutputName,'w');
else
    name=Image_filename{filenum};
    Image_filepath = fullfile(Image_pathname, Image_filename{filenum});
    OutputName=strcat(erase(Image_filename{filenum},[".png",".bmp"]),".txt");
    OutputName=fullfile(Image_pathname, OutputName);
    fileID = fopen(OutputName,'w');
end
II=imread(Image_filepath);
if size(II,3)==3 % 降成2D
    II= rgb2gray(imread(Image_filepath));
end
[verPixel,horPixel]=size(II);
%% 0轉1
if ZerotoOne==1
    II(II==0)=1;
end
%% 開始寫
% 先寫第一排 "字串" EX: MESH: 3840 2160 -60.48 -34.02 60.48 34.02
OutputStr_first=strcat("MESH: ",num2str(horPixel)," ",num2str(verPixel)," ",num2str(-HorSize*0.5)," ",num2str(-VerSize*0.5)," ",num2str(HorSize*0.5)," ",num2str(VerSize*0.5),"\n");
fprintf(fileID,OutputStr_first);
% 寫之後每一行
[mrows, ncols] = size(II); 
outputstr = ['%d ']; % template fo  r the string, you put your datatype here
outputstr = repmat(outputstr, 1, ncols); % replicate it to match the number of columns
outputstr = [outputstr '\n']; %#ok<AGROW> % add a new line if you want
fprintf(fileID,outputstr, II.'); % write it
fclose("all");
disp(strcat(name," 已完成空間切趾"));
end
toc