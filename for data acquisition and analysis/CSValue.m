%% Original Author: BM %%
% 20220620: Modified by GY (project since 20220620)
% �γ~: 13 �I��]�P���R
% �T�O Dummy Plane �M Receiver ���R�W (DX RX)
% �T�O ���B�u���@�� �ҥ�+���u�i�l��ʶ}�� �� CUBE_SURFACE_SOURCE
% Last Update: 20230320
% content: rewrite version
% 1. seedPool
% 2. erosion --> CS
% 3. CS <mean>
close all;clear;clc;
tStart = tic;
%% �ϥΪ̿�J
% Receiver �s��: [1 5 7 8 9 12 13 14 17 18 19 21 25]
% Sheetname �п�u R1 R5 R...
inputFile = "Result_VVA30_HVA0_WDR400_MRM20_03-20-2023 16-52.xlsx";
chooseOutputFolder = 1;
    outputFolderName = "test";

% Erosion Setup
erosionCS= 1;
structureElementRadius = 3;                         % 1 mm �N��X�� Pixel (TBD)
structureElementHeight = 0;                         % offset


%% PreProcessing
targetDirc = outputFolderName;  %�۩w�q��Ƨ��W��
if chooseOutputFolder == 0 % �bMatlab��e��Ƨ��إ� 13�I�Ž���Ƨ�
    rootDirc = string(cd);
elseif chooseOutputFolder == 1 % �u��@����Ƨ���l
    rootDirc = uigetdir;
    if rootDirc == 0;return;end
else
    beep
    error("wrong setup for 'chooseFolder'. should be either 0 or 1.")
end

lastwarn(''); % ���m warning
fullDirc = fullfile(rootDirc,targetDirc);
mkdir(fullDirc);
msg = lastwarn;
if isequal('Directory already exists.',msg)
    disp("< �۩w�q��Ƨ��w�s�b��e�ؿ� >")
end

% excelFilepath = fullfile(fullDirc,inputFile);  % �s���G ��m
excelFilepath = inputFile;

% erosion unit
se = offsetstrel('ball',structureElementRadius, structureElementHeight);

%% read excel
Wr = [1 5 7 8 9 12 13 14 17 18 19 21 25];
Wr_delete = [];

% check Receiver Num
for ww = 1:length(Wr)
    sheetName = strcat("R",num2str(Wr(ww)));
    try
        T = readmatrix(inputFile,'Sheet',sheetName,'Range','A1');
    catch
        Wr_delete = [Wr_delete,ww]; %#ok<AGROW> 
    end
end

% read data
Wr(Wr_delete) = [];
dataArrayAll = cell(length(Wr),1);
for ww = 1:length(Wr)
    sheetName = strcat("R",num2str(Wr(ww)));
    dataTemp = readmatrix(inputFile,'Sheet',sheetName);
    dataArrayAll{ww} = dataTemp(2:end,2:end);
end
Ydim = size(dataArrayAll{1},1);
Xdim = size(dataArrayAll{1},2);
%% ��]�P���R
while (1)
cprintf('key',"��]�P���R......")
LTWDR = zeros(5,10);
intensityArray = zeros(2,13);

Datalist = zeros(5,13);
Datalist(1,:) = Wr;
F_each = cell(13,1);
maxIntensityArray = zeros(5,5);
for temptemp = 1:13 % ��l��
    F_each{temptemp} = zeros(Ydim,Xdim);
end
picTargetErodeArray = cell(5,5);

for k = 1:length(Wr)

    temp = dataArrayAll{k};
    I = temp;
    Datalist(2,k)=roundn(max(temp(:)),-2);
    Datalist(3,k)=roundn(min(temp(:)),-2);
    Datalist(4,k)=roundn(min(temp(:))/max(temp(:)),-2);
    Datalist(5,k)=roundn((min(temp(:))+max(temp(:)))/(max(temp(:))-min(temp(:))),-2);
    if k == 1
        Picc = size(temp);
        Pic = zeros(Picc(1)*5);
        picTargetErodeArray(:,:) = {uint8(zeros(Picc(1),Picc(2)))};
    end
    XX = uint8((temp./max(temp(:)))*255);
    Xmin= double(min(temp(:)))/double(max(temp(:)));

    F_each{k,1}=XX;

    %% Erosion
    if erosionCS == 1
        [row,column] = ind2sub([5,5],Wr(k));
        XX = imerode(XX,se);
        imageFilepath = fullfile(fullDirc,strcat("R",num2str(Wr(k)),"_Erosion.png"));   % �s���G ��m
        imwrite(XX,imageFilepath);
        picTargetErodeArray{row,column} = XX;
    end

    %% �ֿn��Ƨ�10% & 90%
    H_forobserve=histogram(XX);
    H=histogram(XX,'Normalization','cdf');
    edge=[H.BinLimits(1):0.5:H.BinLimits(end)];
    H=histogram(XX,'Normalization','cdf','BinEdges',edge);
    CumData_Y=H.Values;
    CumData_X=zeros(1,length(H.Values));
    for jj=1:length(H.BinEdges)-1
        CumData_X(jj)=0.5*(H.BinEdges(jj)+H.BinEdges(jj+1));
    end
    x_const=0:5:255;  %sim �p���ȥi��ݽվ� 0.1 �j���i�H5  !!!!!
    y_const=x_const;
    y_const(:)=0.1;
    Imin=polyxpoly(x_const,y_const,CumData_X,CumData_Y);
    y_const(:)=0.9;
    Imax=polyxpoly(x_const,y_const,CumData_X,CumData_Y);
    %%
    CR=(Imax-Imin)/(Imax+Imin);   %��X����
    CSF1=1/CR;
    LTWDR(Wr(k)) = roundn(CSF1,-1);
    intensityArray(1,k)=Imax;
    intensityArray(2,k)=Imin;
    
    W_non0 = LTWDR(LTWDR(:,1:5)>0);
    W_non0_AVG=round(sum(W_non0)/13,1);

    %% Imax Ratio 20220609 By GY
    [row,column] = ind2sub([5,5],Wr(k));
    maxIntensityArray(row,column) = Datalist(2,k);
    if k == 1 % �����Ĥ@�դ����I�j�� (��m13) (��7�I)
        maxIntensitySpecific = maxIntensityArray(row,column);
    end
end
if erosionCS == 1
    disp("CS Value (after erosion)");
elseif erosionCS == 0
    disp("CS Value");
end
disp(LTWDR);
disp(strcat("����: ",num2str(W_non0_AVG)));

maxIntensityArrayNormal2Each = maxIntensityArray/max(maxIntensityArray(:));
maxIntensityArrayNormal2Specific = maxIntensityArray/maxIntensitySpecific;

%% Summary
cprintf('key',"�g�JSummary��......")
F_zero=zeros(Picc(1),Picc(2));
Total_F=[F_each{1},F_zero,F_zero,F_zero,F_each{12};
        F_zero,F_each{3},F_each{6},F_each{9},F_zero;
        F_zero,F_each{4},F_each{7},F_each{10},F_zero;
        F_zero,F_each{5},F_each{8},F_each{11},F_zero;
        F_each{2},F_zero,F_zero,F_zero,F_each{13}];

figure;  imshow(Total_F,[0 255]);
imwrite(Total_F,fullfile(fullDirc,"13�I��.png"));
writematrix(LTWDR,excelFilepath,'Sheet','Summary!!','UseExcel',true,'AutoFitWidth',false);
writematrix(W_non0_AVG,excelFilepath,'Sheet','Summary!!','Range','A7','UseExcel',true,'AutoFitWidth',false);
txt={'Max';'Min';'Min/Max(%)';'CS (100%-0%)'};
writecell(txt,excelFilepath,'Sheet','Summary!!','Range','A12','UseExcel',true,'AutoFitWidth',false);
writematrix(Datalist,excelFilepath,'Sheet','Summary!!','Range','B11','UseExcel',true,'AutoFitWidth',false);
writecell({'�̤j���j normal2self'},excelFilepath,'Sheet','Summary!!','Range','P10','UseExcel',true,'AutoFitWidth',false);
writecell({'�̤j���j normal2all'},excelFilepath,'Sheet','Summary!!','Range','W10','UseExcel',true,'AutoFitWidth',false);
writematrix(maxIntensityArrayNormal2Each,excelFilepath,'Sheet','Summary!!','Range','Q10','UseExcel',true,'AutoFitWidth',false);
writematrix(maxIntensityArrayNormal2Specific,excelFilepath,'Sheet','Summary!!','Range','X10','UseExcel',true,'AutoFitWidth',false);
cprintf('err',"����\n")
break
end

%% ���� Imax Imin (13�I) (For CS<mean>)
if erosionCS == 1
    picTargetErodeCombine = cell2mat(picTargetErodeArray);
    imageFilepath = fullfile(fullDirc,"combination_Erosion.png");   % �s���G ��m
    imwrite(picTargetErodeCombine,imageFilepath);
end
tEnd = toc(tStart);
disp(strcat("process complete: ",num2str(tEnd)," seconds."))

%% Function