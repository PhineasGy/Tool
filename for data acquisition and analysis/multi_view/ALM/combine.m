%% 建立資料 table
view_str = ["V1","V2","V3","V4"];
alm = ALM_Table();
for ii = 1:4
    data = readtable("Result_13點向後角輝_MRM 20_P0.1_"+view_str(ii)+".xlsx",...
        "ReadVariableNames",true,"ReadRowNames",true);
    column = linspace(0.5,359.5,360);
    data.Properties.VariableNames = string(column);
    alm.add(data)
end

%% plot (normalize, degree)
figure;
for dd = 1:4
    L_270 = alm.find(dd,col=269.5);
    L_90 = flipud(alm.find(dd,col=89.5));
    x_range = [-89.5:-0.5,0.5:89.5];
    xlabel("horizontal angle (degree)")
    y_range = [L_270',L_90']; y_range = y_range / max(y_range);
    ylabel("luminance (normalized)")
    plot(x_range,y_range);
    hold on 
end
legend(["V1","V2","V3","V4"])

%% plot (normalize, length)
VDZ = 500;
figure;
for dd = 1:4
    L_270 = alm.find(dd,col=269.5);
    L_90 = flipud(alm.find(dd,col=89.5));
    x_range = [-89.5:-0.5,0.5:89.5]; x_range = VDZ * tand(x_range); % 轉換成長度單位
    xlabel("horizontal movement (mm)")
    y_range = [L_270',L_90']; y_range = y_range / max(y_range);
    ylabel("luminance (normalized)")
    plot(x_range,y_range);
    xlim([-240, 240])
    hold on 
end
legend(["V1","V2","V3","V4"])