%% B + C Alone (softII version)
% version: v1.20
% usage: �p�� B + C �`�Ϋ� (�w���Ϫ���)
% update: check fail switch, regional check
clear
clc
% throw(MException("MATLAB:test",""))
tic
%% �ϥΪ̿�J
mask_mode = 1;  % regional VZA
    panel_pixel_number_hor = 2160;
    panel_pixel_number_ver = 3840;

II_mode = "soft";  % soft: 1, hard: 2

is_creating_table = 1;                                        % �O�_�Ыس̫᪺ table (�нT�O VVA HVA �P�ɱҰʥB�ŦX�W�h)
VA_mode = 1:2;                                            % HVA --> VVA�C can be 1, 2, 1:2
WD_list = [400 500 600 700];
VVA_center = 30;
HVA_center = 0;
VVA_list = [5:1:30,30:1:60];              % if createTable == 1, �нT�O center ���ץX�{�⦸�Cex: [-30:0,0:30]
HVA_list = [-40:1:0,0:1:40];              % if createTable == 1, �нT�O center ���ץX�{�⦸�Cex: [-30:0,0:30]

% check TIR
is_checking_TIR = 0;

% check fail (or output all fail number)
is_checking_critical = 0;
    critical_C_fail = 1000;                                 % �ˬd C �� fail number
    critical_B_black = 1000;                                % �ˬd B ���¼ƶq (������)
    critical_BC_fail = 1000;                                % �ˬd B + C �� fail number

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Object Build
CRITICAL = Critical(critical_C_fail=critical_C_fail,critical_B_black=critical_B_black,critical_BC_fail=critical_BC_fail);
MASK = Mask(panel_pixel_number_hor=panel_pixel_number_hor,panel_pixel_number_ver=panel_pixel_number_ver);
OPTION = Option(VA_mode=VA_mode,mask_mode=mask_mode,II_mode=II_mode,...
        is_creating_table=is_creating_table,is_checking_TIR=is_checking_TIR,is_checking_critical=is_checking_critical);
INPUT = Input(VVA_center=VVA_center,HVA_center=HVA_center,...
        WD_list=WD_list,VVA_list=VVA_list,HVA_list=HVA_list,OPTION=OPTION);

%% Main Program
BC_ALONE = BCAlone(INPUT=INPUT,OPTION=OPTION,CRITICAL=CRITICAL,MASK=MASK);
BC_ALONE.open_table;