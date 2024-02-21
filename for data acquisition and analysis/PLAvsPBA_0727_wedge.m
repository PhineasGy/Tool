%% Find Relation between PBA and PLA
close all; 
clc;clear; 
%%
% update: 20230727
% content: wedge prism, direct method only
% PLA: PLA + wedgePrism_PBA (與地面法線夾角)
% output: PLA_array PBA_array_formPLA

%% Input
% 垂直光入射
PBA_check = 37.5;               % For accuracy check: see "PLA_check" and do polyval(p,PLA_check)
n_prism = 1.49;       % 兩者關係只和折射率有關
PBA_array = -90:0.01:90;        % 下層 Prism, [-90:90]    "PBA > 0: 垂直面向人眼"
wedgePrism_PBA = -5;            % 上層 Prism,             "PBA < 0: 垂直面向人眼"
%% Processing
% derive PLA from PBA
PLATemp = PBA_array - asind(sind(PBA_array)/n_prism);
PLA_array = asind(n_prism * sind(PLATemp - wedgePrism_PBA));
% filt TIR Term
PBA_array(imag(PLA_array)~=0) = [];
PLA_array(imag(PLA_array)~=0) = [];
% Plot
figure
plot(PLA_array + wedgePrism_PBA,PBA_array)
xlabel("PLA 出光角");ylabel("PBA 底角")
xlim([-90 90])
xlim([-90 90])
yline(0)
legend(["Data Points"])
PLA_array_fromPBA = PLA_array + wedgePrism_PBA; % 與地面夾角
PBA_array_forPLA = PBA_array;
%% output
disp("output: PLA_array_fromPBA PBA_array_forPLA")
save(strcat("PLAvsPBA",datestr(now,'mm-dd-yyyy HH-MM'),".mat"),"PLA_array_fromPBA","PBA_array_forPLA");
%% interpolation 
% PBA_desired = interp1(PLA_array_fromPBA,PBA_array_forPLA,40);
PLA_desired = interp1(PBA_array_forPLA,PLA_array_fromPBA,PBA_check);
disp(strcat("中心 PLA (GP-PBA: ",num2str(PBA_check),"): ",num2str(PLA_desired)))