%% B + C Alone 
% 20240103
% version: V1.0
% �γ~: �p�� B + C �`�Ϋ� (��J�Ϫ���) (�ثe�Ȥ䴩�n�� II)
clear
clc

%% �ϥΪ̿�J
soft_hard = 1;  % soft: 1, hard: 2

createTable = 1;                                        % �O�_�Ыس̫᪺ table (�нT�O VVA HVA �P�ɱҰʥB�ŦX�W�h)
VANum = 1:2;                                            % HVA --> VVA�C can be 1, 2, 1:2
WD_list = [400];
VVA_center = 30;
HVA_center = 0;
VVA_list_when_VANum2 = [10:1:30,30:1:50];             % if createTable == 1, �нT�O center ���ץX�{�⦸�Cex: [-30:0,0:30]
HVA_list_when_VANum1 = [-30:1:0,0:1:30];              % if createTable == 1, �нT�O center ���ץX�{�⦸�Cex: [-30:0,0:30]
C_fail_critical = 1000;                                 % �ˬd C �� fail number
B_black_critical = 1000;                                   % �ˬd B ���¼ƶq (������)
B_and_C_fail_critical = 1000;                           % �ˬd B + C �� fail number

%% ---------------------------------------------- %%
%% ��ܸ�Ƨ�
currentWorkspace = cd;
% 1. B ��Ƨ�
B_dir_str = uigetdir(cd,"��� B ��Ƨ�");
cd(B_dir_str)
B_png = dir("**/*.png");    % ����]�t�l�ؿ��� png �ɮ�
cd(currentWorkspace)

% % 2. C ��Ƨ�
C_dir_str = uigetdir(cd,"��� C ��Ƨ�");
cd(C_dir_str)
C_png = dir("**/*.png");    % ����]�t�l�ؿ��� png �ɮ�
cd(currentWorkspace)

%% VA Loop 
WDNum = length(WD_list);
dataCell = cell(WDNum,length(VANum));
for wdwd = 1:WDNum
WD = WD_list(wdwd);
for jj = VANum
switch jj
    case 1    
        VVA_list = linspace(VVA_center,VVA_center,length(HVA_list_when_VANum1)); 
        HVA_list = HVA_list_when_VANum1;   
        HVAColumn = HVA_list';
        C_fail_list = nan(1,length(HVA_list));
        BandC_fail_list = nan(1,length(HVA_list));
    case 2
        VVA_list = VVA_list_when_VANum2;                  
        HVA_list = linspace(HVA_center,HVA_center,length(VVA_list_when_VANum2));
        VVAColumn = VVA_list';
        C_fail_list = nan(1,length(VVA_list));
        BandC_fail_list = nan(1,length(VVA_list));
end

C_fail_count = nan;
B_and_C_fail_count = nan;
VA_count = 0;
for whichVA = 1:length(VVA_list)

%% update fail list
if VA_count > 0
    C_fail_list(VA_count) = C_fail_count;
    BandC_fail_list(VA_count) = B_and_C_fail_count;
    C_fail_count = nan;
    B_and_C_fail_count = nan;
end
VA_count = VA_count + 1;
VVA = VVA_list(VA_count);
HVA = HVA_list(VA_count);

%% ��v���A�ˬd TIR
% VA pattern
switch soft_hard
    case 1 % soft
       % EX:
       % [BLP]_IPD=60.00_VD=0400.00_VVA=10.00_HVA=+00.00_PR=41.00_TIR=True.png
        VA_pattern = strcat("_VD=",num2str(WD,"%07.2f"),...
                            "_VVA=",num2str(VVA,"%05.2f"),...
                            "_HVA=",num2str(HVA,"%+06.2f"));
        TIR_check_pattern_before = "TIR=";
        TIR_check_pattern_after = ".";
    case 2 % hard
        % ignore
end

cprintf([1,0.5,0],strcat("[info]: �ؼмv�� ",VA_pattern,"\n"))

B_png_name = string({B_png.name});
B_png_folder = string({B_png.folder});
target_B = B_png_name(contains(B_png_name,VA_pattern));
target_B_folder = B_png_folder(contains(B_png_name,VA_pattern));

if length(target_B) ~= 1
    if VVA == VVA_center && HVA == HVA_center % ��� OOVA �v��: �q�L
        target_B = target_B(1);
        target_B_folder = target_B_folder(1);
    else
        cprintf('err',strcat("[���~]: ������ B �v���ƶq��: ",num2str(length(target_B)),"\n"))
        cprintf('err',"�t�ΰ���\n")
        beep
        return
    end
end
C_png_name = string({C_png.name});
C_png_folder = string({C_png.folder});
target_C = C_png_name(contains(C_png_name,VA_pattern));
target_C_folder = C_png_folder(contains(C_png_name,VA_pattern));

if length(target_C) ~= 1
    if VVA == VVA_center && HVA == HVA_center % ��� OOVA �v��: �q�L
        target_C = target_C(1);
        target_C_folder = target_C_folder(1);
    else
        cprintf('err',strcat("[���~]: ������ C �v���ƶq��: ",num2str(length(target_C)),"\n"))
        cprintf('err',"�t�ΰ���\n")
        beep
        return
    end
end

% TIR �ˬd
TIR_boolean_B = extractBetween(target_B,TIR_check_pattern_before,TIR_check_pattern_after);
TIR_boolean_C = extractBetween(target_C,TIR_check_pattern_before,TIR_check_pattern_after);

switch TIR_boolean_B
    case "True"
        cprintf('key',"[info]: B �o�� TIR (continue to next)\n")
        continue
    case "False"
        % pass
    otherwise
        cprintf('err',"[���~]: �L�k���� B �� TIR ����\n")
        cprintf('err',"�t�ΰ���\n")
        beep
        return
end
switch TIR_boolean_C
    case "True"
        cprintf('key',"[info]: C �o�� TIR (continue to next)\n")
        continue
    case "False"
        % pass
    otherwise
        cprintf('err',"[���~]: �L�k���� C �� TIR ����\n")
        cprintf('err',"�t�ΰ���\n")
        beep
        return
end
% imread
fullpath_B = fullfile(target_B_folder,target_B);
fullpath_C = fullfile(target_C_folder,target_C);
B = double(imread(fullpath_B));   % �z�פW�������
C = double(imread(fullpath_C));   % �z�פW�������
% image handle
switch soft_hard
    case 1
        maxRGB_B = max(B,[],"all");
        B = B * 0.5;
        maxRGB_C = max(C,[],"all");
        C = C * 0.5;
        if maxRGB_B ~= maxRGB_C
            error(strcat("B and C RGB �̤j�Ȥ��ۦP: B:", num2str(maxRGB_B)," C:",num2str(maxRGB_C)));
        end
    case 2
        % ignore
end

%% C Fail check
[C_fail_array,~] = find(C(:,:,1) >= 0.5 * maxRGB_C & C(:,:,2) >= 0.5 * maxRGB_C);
C_fail_count = length(C_fail_array);
disp(strcat("[info]: C fail number: ",num2str(C_fail_count)));
if C_fail_critical < C_fail_count
    cprintf('key',"[result]: C Fail (���L�h) (continue to next)\n")
    continue
end

%% B �� check
[B_black_array,~] = find(B(:,:,1) == 0 & B(:,:,2) == 0 & B(:,:,3) == 0);
B_black_count = length(B_black_array);
disp(strcat("[info]: B black number: ",num2str(B_black_count)));
if B_black_critical < B_black_count
    cprintf('key',"[result]: B Fail (�¹L�h) (continue to next)\n")
    continue
end

%% B + C Fail check
B_C = B + C;
[B_and_C_fail_array,~] = find(B_C(:,:,1) >= 0.5 * maxRGB_C & B_C(:,:,2) >= 0.5 * maxRGB_C);
B_and_C_fail_count = length(B_and_C_fail_array);
disp(strcat("[info]: B + C fail number: ",num2str(B_and_C_fail_count)));
if B_and_C_fail_critical < B_and_C_fail_count
    cprintf('key',"[result]: B + C Fail (���L�h) (continue to next)\n")
    continue
end

%% Pass: 
cprintf('key',"[result]: Pass\n")
C_fail_list(VA_count) = C_fail_count;
BandC_fail_list(VA_count) = B_and_C_fail_count;
end % VA Loop
%% update fail list (final but fail term)
if VA_count == length(C_fail_list)
    C_fail_list(VA_count) = C_fail_count;
    BandC_fail_list(VA_count) = B_and_C_fail_count;
end

%% ���R���� ��z���
switch jj
    case 1
        dataCell{wdwd,jj}.C_fail_list = C_fail_list;
        dataCell{wdwd,jj}.BandC_fail_list = BandC_fail_list;
    case 2
        dataCell{wdwd,jj}.C_fail_list = C_fail_list';
        dataCell{wdwd,jj}.BandC_fail_list = BandC_fail_list';
end
end % VANum
end % VD
cprintf("=== ���R���� ===\n")

%% �Ы� final table
if createTable == 1
    finalTable(length(WD_list)).BC = nan;
    
    tableBase = cell(length(VVAColumn)+1,length(HVAColumn)+1);
    tableBase(:) = {""};
    tableBase(1,1) = {"VVA/HVA"};
    tableBase(2:end,1) = num2cell(VVAColumn);
    tableBase(1,2:end) = num2cell(HVAColumn);
    
    for whichWD = 1:length(WD_list)
        
        tableBC = tableBase;
        tableBC(1+find(VVAColumn == VVA_center,1,"first"),2:end) = num2cell(dataCell{whichWD,1}.C_fail_list);
        tableBC(1+find(VVAColumn == VVA_center,1,"last"),2:end) = num2cell(dataCell{whichWD,1}.BandC_fail_list);
        tableBC(2:end,1+find(HVAColumn == HVA_center,1,"first")) = num2cell(dataCell{whichWD,2}.C_fail_list);
        tableBC(2:end,1+find(HVAColumn == HVA_center,1,"last")) = num2cell(dataCell{whichWD,2}.BandC_fail_list);
        finalTable(whichWD).BC = tableBC;
    
    end
end