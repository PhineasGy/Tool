function func_spa(II,OutputII_name,HorSize,VerSize,Zero2One)
%% Function: Spatial apodization Matlab version (空間切趾 For LT Sim)　20220125 GY %%
% < func_spa > 
% ----------------------------------------------
% input: (II,HorSize,VerSize,Zero2One)
% 	[
% 		II: 欲轉成切趾檔的II圖
% 		HorSize: 面板長邊長度(mm)
% 		VerSize: 面板短邊長度(mm)
% 		Zero2One: 0轉1設定
% 	]
% 
% output: 
% 	[
% 		spaFile: 切趾檔 (txt文件檔)
% 	]
%
% Usage:
% EX: func_spa(II,120.96,68.04,0)
% ----------------------------------------------
%% User Decide
% Zero2One=0; % 是否將Matrix中的0轉1
% HorSize=120.96;
% VerSize=68.04;
%% 處理 II圖
tic

OutputName=strcat(erase(OutputII_name,[".png",".bmp"]),".txt");
fileID = fopen(OutputName,'w');

if size(II,3)==3 % 降成2D
    II= rgb2gray(II);
end
[verPixel,horPixel]=size(II);
%% 0轉1
if Zero2One==1
    II(II==0)=1;
end
%% 轉換切趾檔
% 第一排 "字串" EX: MESH: 3840 2160 -60.48 -34.02 60.48 34.02
OutputStr_first=strcat("MESH: ",num2str(horPixel)," ",num2str(verPixel)," ",num2str(-HorSize*0.5)," ",num2str(-VerSize*0.5)," ",num2str(HorSize*0.5)," ",num2str(VerSize*0.5),"\n");
fprintf(fileID,OutputStr_first);
% 寫之後每一行
[~, ncols] = size(II); 
outputstr = ['%d ']; %#ok<NBRAK> % template fo  r the string, you put your datatype here
outputstr = repmat(outputstr, 1, ncols); % replicate it to match the number of columns
outputstr = [outputstr '\n']; % add a new line if you want
fprintf(fileID,outputstr, II.'); % write it
fclose("all");
disp("已完成空間切趾");
toc
end