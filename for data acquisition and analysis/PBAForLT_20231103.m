%% LT 漸變 Prism 清單 (wedge version) %%
% author: GY
% starting date: 20231103
% content: wedge 3D
close all; 
clc;clear;
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Input %%
%% LT setup
LTFirstPoint = [-171.8999938964844;0];      % mm LT中Prism漸變位置清單的起點座標 LT:(X,Y)
LTFinalPoint = [171.8000030517578;0];       % mm LT中Prism漸變位置清單的終點座標 LT:(X,Y)

manualLTNum = 0;                            % 是否手動輸入PrismArray數量
    LTNumPrism = 3438;                      % 手動輸入 PrismArray 數量 

% AutoGP 相關參數
prism_n = 1.49;            
prism_pitch = 0.1;
gp_PBA_mid = 41.5;                          % center gp data
gp_PRA = 0;                                 % gp_PRA 保持 0 度 (for LT)
gp_WD_mid = 550;                            % center gp data, from "wedge top"
PSA = 71.02;
wedgePrism = 1;
    wp_verSize = 343.76;                    % for wedge height
    wp_PBA = -5;            
    wp_PRA = 10;                            % 相對 gp_PRA

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% preprocessing %%
air_n = 1;
systemThickness = 0; % 後面計算會消去
rotGPPRA = rotz(gp_PRA);
% LT pointA handle
if manualLTNum == 0
    LTNumPrism = round((LTFinalPoint(1)-LTFirstPoint(1))/prism_pitch + 1); % LT中Prism漸變位置清單之Prism 總數
end
LT_X = linspace(LTFirstPoint(1),LTFinalPoint(1),LTNumPrism);
LT_Y = linspace(LTFirstPoint(2),LTFinalPoint(2),LTNumPrism);
Matlab_X = -LT_X;
Matlab_Y = LT_Y;
pointA_array = [Matlab_X;Matlab_Y];
PBA_array = nan(1,LTNumPrism);

%% processing %%
%% step 1: get virtual eye
% PLA (wedge version) 20230728
if wedgePrism == 0
    wp_PRA = 0;
    wp_PBA = 0;
end
% wedge 法向量
rotWPPRA = rotz(wp_PRA);
wedge_normal = [-sind(wp_PBA);0;-cosd(wp_PBA)];
wedge_normal = rotWPPRA * wedge_normal; % 旋轉法向量
% wedge 厚度(高度): 
wp_height_max = abs(wp_verSize*tand(wp_PBA));
wp_height_half = wp_height_max/2;

% 計算中心 PLA (autoGP)
forPLA_incident = [0;0;1];
temp_normalGP = [sind(gp_PBA_mid);0;cosd(gp_PBA_mid)];
forPLA_normalGP = rotGPPRA * temp_normalGP;
forPLA_normalWP = -wedge_normal;
forPLA_mu_airtoPrism = air_n / prism_n;
forPLA_mu_PrismtoAir = 1/forPLA_mu_airtoPrism;
% tracing
% part 1: air into prism (through GP)
forPLA_insidePrism = sqrt(1-forPLA_mu_airtoPrism^2*(1-(forPLA_normalGP'*forPLA_incident)^2))*forPLA_normalGP+...
                        forPLA_mu_airtoPrism*(forPLA_incident-(forPLA_normalGP'*forPLA_incident)*forPLA_normalGP); %snell's law <air到Lens>         
% part 2: prism into air (through WP)
forPLA_outsidePrism = sqrt(1-forPLA_mu_PrismtoAir^2*(1-(forPLA_normalWP'*forPLA_insidePrism)^2))*forPLA_normalWP+...
                        forPLA_mu_PrismtoAir*(forPLA_insidePrism-(forPLA_normalWP'*forPLA_insidePrism)*forPLA_normalWP); %snell's law <air到Lens> 
% PLA: 與 Z 軸夾角
mid_PLA = acosd(dot(forPLA_outsidePrism,[0,0,1]));
if forPLA_outsidePrism(1) < 0; mid_PLA = -mid_PLA; end

% PLA vs PBA function
gp_PBA_array = -90:0.01:90;
[PLA_array_fromPBA,PBA_array_forPLA] = PLAvsPBA(gp_PBA_array,air_n,prism_n,rotGPPRA,wp_PBA,rotWPPRA);
PLA_midCheck = interp1(PBA_array_forPLA,PLA_array_fromPBA,gp_PBA_mid);
disp(strcat("中心 PLA (GP-PBA: ",num2str(gp_PBA_mid),"): ",num2str(PLA_midCheck)))

VE_originalPoint = [0;0;wp_height_half]; % assume substrate top: Z = 0
VE_EyePoint = VE_originalPoint + forPLA_outsidePrism * (gp_WD_mid*cosd(mid_PLA)/forPLA_outsidePrism(3));

%% step 2: get desired PBA (in loop IPA)
for whichPoint = 1:LTNumPrism
    pointA = nan(3,1);
    pointA(1:2) = pointA_array(:,whichPoint); % substrate top position 

    % update wedge height
    if wedgePrism == 1
        pointXY = [pointA(1);pointA(2)]; % example
        pointXYRot = rotWPPRA \ [pointXY;0]; % 反矩陣 (座標轉換)
        pointXYRot = pointXYRot(1:2);
        pointXYRotShift = pointXYRot + [wp_verSize/2;0];
        wedgeHeight = abs(pointXYRotShift(1)*tand(wp_PBA));
        pointA(3) = wedgeHeight;
    else
        pointA(3) = 0; % no wedge, Z = 0
    end

    if norm(VE_EyePoint(1:2))~=0 
        length_wedgeTop_virtualEyeLine_xy = abs(pointA(1:2)'*VE_EyePoint(1:2)-norm(VE_EyePoint(1:2))^2)/norm(VE_EyePoint(1:2)); % 點到線距離
        % 判斷 L 正負 20230808
        % 假設 零點 (0,0) 一定在虛擬眼之上
        % if L 和 零點 在切線同一側: L > 0
        % 切線方程式: E𝑦∗𝑦+𝐸𝑥∗𝑥−(𝐸𝑦^2+𝐸𝑥^2)=0
        centerSide = -norm(VE_EyePoint(1:2))^2;
        targetSide = pointA(1:2)'*VE_EyePoint(1:2) - norm(VE_EyePoint(1:2))^2;
        if centerSide*targetSide < 0;length_wedgeTop_virtualEyeLine_xy = -length_wedgeTop_virtualEyeLine_xy;end    
    else % 虛擬眼睛與面板中心重疊時
        if gp_PRA == 0
            length_wedgeTop_virtualEyeLine_xy = -pointA(1);
        elseif gp_PRA ==90 || gp_PRA ==-90
            length_wedgeTop_virtualEyeLine_xy = -pointA(2);
        else
            length_wedgeTop_virtualEyeLine_xy = -pointA(1)*cosd(gp_PRA);
        end
    end
    PLA_desired = atand(length_wedgeTop_virtualEyeLine_xy/(VE_EyePoint(3)-pointA(3)));
    PBA_desired = interp1(PLA_array_fromPBA,PBA_array_forPLA,PLA_desired);
    PBA_array(whichPoint) = PBA_desired;
end

%% 輸出變數清單: angleWest angleEast
LT_PBA = PBA_array';
LT_PBA = flipud(LT_PBA);
%% LT west east angle list
% 一般情形:             angle west: PBA, angle east = PSA
% when PBA < 0 -->      angle east = |PBA|, angle west = PSA
if all(LT_PBA>0);angleWest = LT_PBA;angleEast = PSA;end
if any(LT_PBA<0)
    angleWest = LT_PBA;
    angleEast = PSA * ones(LTNumPrism,1);
    angleEast(angleWest<0) = abs(angleWest(angleWest<0));
    angleWest(angleWest<0) = PSA;
end
disp("輸出變數清單: angleWest, angleEast");
plot(LT_PBA)
open angleWest
open angleEast

%% function
function [PLA_array_fromPBA,PBA_array_forPLA] = PLAvsPBA(gp_PBA_array,air_n,prism_n,...
    rotPRA,wp_PBA,rotWedgePRA)
    % 20231027: 三維向量法
    % 因為大部分情形 GP PRA 不等於 Wedge PRA， PLA 必須使用三維計算
    % incident vector: 垂直入射
    forPLA_incident = [0;0;1];
    % normal vector: WP (fixed)
    wedge_normal = [-sind(wp_PBA);0;-cosd(wp_PBA)];
    wedge_normal = rotWedgePRA * wedge_normal;
    forPLA_normalWP = -wedge_normal; % 往 +Z 追跡
    % other parameters
    forPLA_mu_airtoPrism = air_n / prism_n;
    forPLA_mu_PrismtoAir = 1/forPLA_mu_airtoPrism;
    
    % tracing to get PLA
    count = 0;
    PLA_array = nan(1,length(gp_PBA_array));
    for gp_PBA = gp_PBA_array
        count = count + 1;
        % normalGP update
        temp_normalGP = [sind(gp_PBA);0;cosd(gp_PBA)];
        forPLA_normalGP = rotPRA * temp_normalGP;
        % tracing
        % part 1: air into prism (through GP)
        forPLA_insidePrism = sqrt(1-forPLA_mu_airtoPrism^2*(1-(forPLA_normalGP'*forPLA_incident)^2))*forPLA_normalGP+...
                                forPLA_mu_airtoPrism*(forPLA_incident-(forPLA_normalGP'*forPLA_incident)*forPLA_normalGP); %snell's law <air到Lens>         
        % part 2: prism into air (through WP)
        forPLA_outsidePrism = sqrt(1-forPLA_mu_PrismtoAir^2*(1-(forPLA_normalWP'*forPLA_insidePrism)^2))*forPLA_normalWP+...
                                forPLA_mu_PrismtoAir*(forPLA_insidePrism-(forPLA_normalWP'*forPLA_insidePrism)*forPLA_normalWP); %snell's law <air到Lens> 
        % PLA: 與 Z 軸夾角
        PLA = acosd(dot(forPLA_outsidePrism,[0,0,1]));

        % PLA 正負判定:
        % X 向量: forPLA_outsidePrism(1) < 0 --> PLA < 0
        if forPLA_outsidePrism(1) < 0; PLA = -PLA; end
        PLA_array(count) = PLA;
    end
    % filt TIR Term
    gp_PBA_array(imag(PLA_array)~=0) = [];
    PLA_array(imag(PLA_array)~=0) = [];
    PLA_array_fromPBA = PLA_array;
    PBA_array_forPLA = gp_PBA_array;
end