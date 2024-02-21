%% 非球面法向量計算程式 %% author: GY
clear
clc
close all
%%%%%%%%%%%%%%%%%%%%%%%%
% version: 20230721
% 等效球面資訊
%%%%%%%%%%%%%%%%%%%%%%%%
%% input
% mode option
plotCurve = 1;
    samplingNumber = 1000;      % sample number for plotting curve
writeMat = 1;
    matFile = ".mat";           % 輸出檔名: strcat("(asph) ",matfile)

% LL 參數
lensMaxAperture = 1;            % 即 Lens Pitch
% 非球面係數
R = 1.533628709150636;            % curvature (mm)
k = -62.65411936271715;          % konic
a = [0,0,...                    % 多項式係數 (start from 1次)
    0,-3.967492033831446E-001,...         3 4 項
    0,4.885193001453348E+000,...          5 6 項
    0,-1.555176494467397E+001,...         7 8 項
    0,1.958501395369059E+001,...          9 10 項
    0,-1.357443762774419E+000,...         11 12 項
    0,-9.381754667029192E+000,...         13 14 項
    0,5.053512773311853E+000]; 

%%%%%%%%%%%%%%%%%%%%%%%%
%% get normal
syms rad
rad2 = rad.^2;
sag = ( rad2 / R ) ./ ( 1 + sqrt( 1 - ( k + 1 ) * rad2 ./ R^2 ) );
% sag( ~isreal( sag ) ) = 1e+20; % prevent complex values
for i = 1 : length( a )
    sag = sag + a( i ) * rad.^i;
end

endpoint_left = [-lensMaxAperture/2,double(subs(sag,rad,-lensMaxAperture/2))];
endpoint_right = [lensMaxAperture/2,double(subs(sag,rad,lensMaxAperture/2))];

drad = gradient(sag);
dseg = -gradient(rad);
normal = [drad,dseg];
leftEnd_normal = -double(subs(normal,[rad,sag],endpoint_left)); % "-": 指向 Curve 內
rightEnd_normal = -double(subs(normal,[rad,sag],endpoint_right)); % "-": 指向 Curve 內
% for coding (座標轉換)

asph_leftEnd_normal = [0;leftEnd_normal(1);-leftEnd_normal(2)];
asph_leftEnd_normal = asph_leftEnd_normal/norm(asph_leftEnd_normal);
asph_rightEnd_normal = [0;rightEnd_normal(1);-rightEnd_normal(2)];
asph_rightEnd_normal = asph_rightEnd_normal/norm(asph_rightEnd_normal);
% 左右對稱


%% get lens max height
% abs(sag(lensAperture/2) - sag(0))
asph_max_height = double(abs(subs(sag,rad,lensMaxAperture/2)-subs(sag,rad,0)));

%% plot curve
if plotCurve == 1
    rad_plot = linspace(-lensMaxAperture/2,lensMaxAperture/2,samplingNumber);
    seg_plot = double(subs(sag,rad,rad_plot));
    plot(rad_plot,seg_plot)
    xlabel("r (lens aperture)")
    axis equal
    % 驗算
    direction = [rad_plot(1)-rad_plot(2),seg_plot(1)-seg_plot(2)];
    disp(strcat("端點法線與方向向量內積值: ",num2str(dot(direction,leftEnd_normal))))
end

%% write mat
if writeMat == 1
    save(strcat("(asph) ",matFile),'asph_leftEnd_normal','asph_rightEnd_normal','asph_max_height')
end

%% 等效球面 R
asph_leftEnd_normal;
theta = abs(atand(asph_leftEnd_normal(2)/asph_leftEnd_normal(3)));
radius_sphere = lensMaxAperture/2/sind(theta);
sag_sphere = radius_sphere - sqrt(radius_sphere^2-(0.5*lensMaxAperture)^2);
disp("< 等效球面資訊 >")
disp(strcat("球半徑 (mm): ",num2str(radius_sphere)))
disp(strcat("球高 (mm): ",num2str(sag_sphere)))