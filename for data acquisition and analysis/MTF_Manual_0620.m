clc;clear;close all
%% 使用者選項
cropTime = 2; % 抓取範圍次數
horVer = 1; % -1:橫 1:直
writeResult = 0; % 是否寫出結果圖

%% Processing
%% 讀圖 (multiselect support)
Image_pathname=[];
[Image_filename, Image_pathname] = uigetfile({strcat(Image_pathname,'*.png;',Image_pathname,'*.jpg')}, '原圖','MultiSelect', 'on');
if ~ischar(Image_pathname) 
    return;end
if ischar(Image_filename)
    totalnum_file=1;
else
    totalnum_file=length(Image_filename);
end

%% Loop 前 參數設定
horVerStr = ["橫條紋","直條紋"];
horVerStr2 = ["位置(上至下)","位置(左至右)"];
disp(strcat("處理模式: ",horVerStr(uint8(horVer/2+1.5))));

for filenum = 1:totalnum_file
    if totalnum_file == 1
        Image_filepath = fullfile(Image_pathname, Image_filename);
        ImageName = Image_filename;
    else
        Image_filepath = fullfile(Image_pathname, Image_filename{filenum});
        ImageName = Image_filename{filenum};
    end
    grayOriginal0=imread(Image_filepath);
    grayOriginal=imrotate(grayOriginal0,0);
    %% 前處理
    doItAgain = 1;
    isRecord = 0;
    while (doItAgain == 1)
        close all
        % Imcrop 手動 (N次)
        grayCrop = cell(1,cropTime);
        grayBefore = grayOriginal;
        restart1 = 0;
        for ct = 1:cropTime
            fh = figure("Name",ImageName);
            fh.WindowState = 'maximized';
            [grayCrop{ct},~] = imcrop(grayBefore); 
            if isempty(grayCrop{ct}) % 沒有切圖
                doItAgain = AskMeQuit();
                if doItAgain == 0 % 不想再切 --> 紀錄 --> 換下一張圖
                    break
                elseif doItAgain == 1
                    restart1 = 1; % 繼續切
                    break
                elseif doItAgain == -1 % 不想再切 --> 系統停止
                    return
                end
                
            end
            grayBefore = grayCrop{ct};
        end
        if restart1 == 1 % 回到 while loop 開頭
            continue
        end
        if doItAgain == 0 % 離開 while(doitagain) 紀錄完 換下一張圖
            isRecord = 1;
            break
        end

        grayFinal = grayCrop{end};
        
        % 24 位元 --> 8 位元
        if size(grayFinal,3) == 3
            grayFinal=rgb2gray(grayFinal);    
        end
        % 擷取一段資訊分布 (均值)
        switch horVer
            case -1 % 橫
                result = mean(grayFinal,2)'; %1直 2橫
                order = 1;
            case 1 % 直
                result = mean(grayFinal,1); %1直 2橫
                order = 2;
        end
        % 寫出crop後的影像
        imshow(grayFinal);
        if writeResult == 1
            imwrite(grayFinal,strcat(horVerStr(order),"_result_File",num2str(filenum),".png"));
        end
        % 紀錄該分布
        fig=figure;
        plot(result,"LineWidth",3,"Color","k");
        title("Intensity分布");xlabel(horVerStr2(order));ylabel("灰階(0-255)");    
        
        %% MTF 解析判定及預估
        %% step 1: 找峰值 谷值
        totalRange = length(result);
        [value_pks, loc_pks] = findpeaks(result); % 找峰值
        [value_valleys, loc_valleys] = findpeaks(-result); % 找谷值 (谷值為負) value_valleys<0
        value_valleys = -value_valleys; % 讓value_valleys>0
    
        hold on
        plot(loc_pks,  value_pks, '^r')
        plot(loc_valleys, value_valleys, 'vg')
    
        %% step 2: 用閥值消除以下的Peak
        thresholdValue = mean(result);
        value_pks_step2 = value_pks(value_pks>thresholdValue);
        loc_pks_step2 = loc_pks(value_pks>thresholdValue);
    
        %% step 3: 去除沒有在兩Peak之間的Valley (消兩側的VALLEY)
        correct_valley = find(loc_valleys > min(loc_pks_step2) & loc_valleys < max(loc_pks_step2));
        value_valleys_step3 = value_valleys(correct_valley);
        loc_valleys_step3 = loc_valleys(correct_valley);
    
        % 不可解析情形1
        if isempty(value_valleys_step3) % 已經沒有 Valley
            MTF_output = 0;
            disp("去除兩Peak外的Valley後, 已無有效Valley : 不可解析")
            if writeResult == 1
                print(fig,fullfile(cd,... 
                    strcat(horVerStr(order),'_MTF無法解析(去除兩Peak外的Valley無有效Valley_失敗)_File',num2str(filenum),'.png')),'-dpng');
            end
            doItAgain = AskMeQuit();
            switch doItAgain
                case 1 % 繼續 Loop
                    continue
                case 0 % 離開 while(doitagain) 紀錄完 換下一張圖
                    isRecord = 1;
                    break
                case -1 % 系統停止
                    return
            end 
            
        end
    
        %% step 4: 利用最小Valley先切邊
        % 頭尾切除法
        % 切除三波以外區域
        % 若沒有取到頭尾 -- 不切
        step5_non = 0;
        x = 1:totalRange;
        y = result; %直條紋
    
        [y_seg,~] = min(value_valleys_step3); % 取最小谷值當作橫切線 
        aa = value_valleys_step3 - y_seg;
        bb = find(~aa); % 找 aa==0 的位置
        x_const = 0:0.1:totalRange;
        y_const = zeros(1,length(x_const));
        
    %     y_const(:) = y(loc_valleys_step3(min(bb))); % 避免有多個位置同時有最小值 
    
        % 特殊切法:考量最低谷值和y(loc_valleys_step3(min(bb)))的權重 (提高SineFitting成功率)
        clip_point = (min(value_valleys)*0.2+y(loc_valleys_step3(min(bb)))*0.8);
        y_const(:) = clip_point; % 避免有多個位置同時有最小值 
        
        % 交點X值
        intersection_points=polyxpoly(x_const,y_const,x,y); % 預期頭尾要各有一個點
        
        ii=find(intersection_points<min(loc_pks_step2),1,'last'); %index % 小於minPeak的最大交點
        jj=find(intersection_points>max(loc_pks_step2),1,'first'); %index % 大於maxPeak的最小交點
        if ~isempty(ii)
            B=intersection_points(ii);
        elseif isempty(ii)
            ii=0;
            B=x(1); % 前端沒有交點: x(1)當第一個值
            step5_non=1; % 沒切到頭尾的 不做step5
        end
        if ~isempty(jj)
            C=intersection_points(jj);
        elseif isempty(jj)
            jj=size(intersection_points,2)+1; 
            C=x(end); % 後端沒有交點: x(end)當第一個值
            step5_non=1; % 沒切到頭尾的 不做step5
        end
        intersection_points=[B;intersection_points(ii+1:jj-1);C]; % 新範圍
        y = y(floor(min(intersection_points)):ceil(max(intersection_points))); % 新Y (取整
        x = 1:length(y); % 新X (取整 ("位移到1"
    
        % 位移 峰值谷值
        loc_valleys_step4=loc_valleys_step3-floor(min(intersection_points))+1;
        loc_pks_step4=loc_pks_step2-floor(min(intersection_points))+1;
    
        %% step 5: 形態學 消除錯誤位置的Peak 和 Valley (不採用)
        loc_valleys_step5=loc_valleys_step4;value_valleys_step5=value_valleys_step3;
        loc_pks_step5=loc_pks_step4;value_pks_step5=value_pks_step2;
        eachpart=(max(intersection_points)-min(intersection_points))/12; % 
    
        %% step 6: 最小Peak*0.8=critical valley value: 兩側Valley判定
        % 取最小的峰值(>T),乘上0.8(Rayleigh Criterion) 得到最大允許谷值V <VinBothSide Method>
        critical_valley = zeros(2,1);
        restart2 = 0;
        for i = [1,5] %(分左右邊判定)  % 左: 第1,2波 右: 第2,3波
            ii = i*0.25+0.75;
            value_pks_step5_half=value_pks_step5((1+eachpart*i)<loc_pks_step5 & loc_pks_step5<(1+eachpart*(i+6)));
            value_valleys_step5_half=value_valleys_step5((1+eachpart*i)<loc_valleys_step5 & loc_valleys_step5<(1+eachpart*(i+6)));
            [min_pk,~]=min(value_pks_step5_half);
            
            % 不可解析情形2
    
            if isempty(min_pk) % 某半邊完全沒有 PK or VL
                MTF_output=0;
                disp("某半邊無Valley or Peak : 不可解析")
                fig=figure;plot(x,y);title("Intensity分布");xlabel(horVerStr2(order));ylabel("灰階(0-255)"); hold on
                plot(loc_pks_step5,  value_pks_step5, '^r')
                plot(loc_valleys_step5, value_valleys_step5, 'vg')    
                line([1+eachpart*1,1+eachpart*7], [critical_valley(1),critical_valley(1)],'Color','black');
                line([1+eachpart*5,1+eachpart*11], [critical_valley(2),critical_valley(2)],'Color','black');
                if writeResult == 1
                    print(fig,fullfile(cd,... 
                        strcat(horVerStr(order),'_MTF無法解析(某半邊無Valley or Peak)_File',num2str(filenum),'.png')),'-dpng');
                end
                doItAgain = AskMeQuit();
                switch doItAgain
                    case 1 % 繼續 Loop
                        restart2 = 1;
                        break
                    case 0 % 離開 while(doitagain) 紀錄完 換下一張圖
                        break
                    case -1 % 系統停止
                        return
                end 
            end
    
            critical_valley(ii) = min_pk*0.8;
    
            % 不可解析情形3
            if ~any(value_valleys_step5_half<min_pk*0.8) % 小於 CVV 沒有任何 谷值: 不可解析
                MTF_output=0;
                disp("一分為2法CVV 判定失敗: 不可解析")                
                fig=figure;plot(x,y);title("Intensity分布");xlabel(horVerStr2(order));ylabel("灰階(0-255)");hold on
                plot(loc_pks_step5,  value_pks_step5, '^r')
                plot(loc_valleys_step5, value_valleys_step5, 'vg')     
                line([1+eachpart*1,1+eachpart*7], [critical_valley(1),critical_valley(1)],'Color','black');
                line([1+eachpart*5,1+eachpart*11], [critical_valley(2),critical_valley(2)],'Color','black');
                if writeResult == 1
                    print(fig,fullfile(cd,...
                        strcat(horVerStr(order),'_MTF無法解析(一分為2法CVV 判定失敗)_File',num2str(filenum),'.png')),'-dpng');
                end
                doItAgain = AskMeQuit();
                switch doItAgain
                    case 1 % 繼續 Loop
                        restart2 = 1;
                        break
                    case 0 % 離開 while(doitagain) 紀錄完 換下一張圖
                        break
                    case -1 % 系統停止
                        return
                end 
                
            end 
        end
        if restart2 == 1 % 繼續 Loop
            continue
        end
        if doItAgain == 0 % 離開 while(doitagain) 紀錄完 換下一張圖
            isRecord = 1;
            break
        end
    
        %% 解析判定end
        disp("恭喜過關: 進入MTF計算...")
    
        %% MTF 演算 (Sine Fitting) %% 已過解析判斷
        try
        [SineP]=sineFit(x,y,0); %Offset, amplitude, frquency, phase, MSE
        catch e % 找不到最佳頻率 (可能區間會有問題 先暫時不算該MTF)
            if (strcmp(e.identifier,'MATLAB:UndefinedFunction'))
                beep
                errordlg("錯誤: 找不到 'sineFit' Matlab Function檔案")
                return
            else
                MTF_output=0;
                disp("SineFitting找不到最佳解(頻率): 暫不可解析")     
                if writeResult == 1
                    fig=figure;plot(x,y);title("Intensity");xlabel(horVerStr2(order));ylabel("輝度nits");
                    print(fig,fullfile(cd,...
                        strcat(horVerStr(order),'_MTF暫無法解析(SineFitting找不到最佳解(頻率))_File',num2str(filenum),'.png')),'-dpng');
                end
                doItAgain = AskMeQuit();
                switch doItAgain
                    case 1 % 繼續 Loop
                        continue
                    case 0 % 離開 while(doitagain) 紀錄完 換下一張圖
                        isRecord = 1;
                        break
                    case -1 % 系統停止
                    return
                end 
            end
        end
        
        % 0: 不秀圖 (FFT)
        s=[SineP(2) 1/SineP(3) SineP(4) SineP(1)];
        xp = linspace(min(x),max(x));
        % disp(s(1));  % s(1): sine wave amplitude (in units of y)
        % disp(s(2));  % s(2): period (in units of x)
        % disp(s(3));  % s(3): phase (phase is s(2)/(2*s(3)) in units of x)
        % disp(s(4));  % s(4): offset (in units of y)
        
        % 平移到最高Peak
        s(4)=max(value_pks_step5)-s(1);
        % 低估Model
        y_fit = s(4) + s(1) * sin(2*pi *1/s(2) * xp +s(3));
        
        Imax=s(4)+s(1); % offset+amplitude
        Imin=s(4)-s(1); % offset-amplitude
        
        MTF_value = (Imax-Imin)/(Imax+Imin); % output MTF if has
        MTF_value(MTF_value>1)=1;
        MTF_value(MTF_value<0)=0;
    
        %% Intensity Plot  %% 已過解析判斷
        fig=figure;
        plot(x,y,'b',  xp,y_fit, 'r')
        hold on;
        plot(loc_pks_step5,  value_pks_step5, '^r')
        plot(loc_valleys_step5, value_valleys_step5, 'vg')       
        title(strcat("MTF:",num2str(MTF_value)));xlabel(horVerStr2(order));ylabel("灰階(0-255)");
        xlim([1 length(y)]);
        grid
        yline(thresholdValue,'-','Threshold Value (from Binarization)');
        line([1+eachpart*1,1+eachpart*7], [critical_valley(1),critical_valley(1)],'Color','black');
        line([1+eachpart*5,1+eachpart*11], [critical_valley(2),critical_valley(2)],'Color','black');
        if writeResult == 1
            print(fig,fullfile(cd,strcat(horVerStr(order),'_MTF_',num2str(MTF_value),"_File",num2str(filenum),'.png')),'-dpng');
        end
        MTF_output = MTF_value; % 輸出 橫/直 條紋的 MTF
        disp(strcat("MTF值為: ",num2str(MTF_output)));

        %% 詢問是否再做一次
        doItAgain = AskMeQuit();
        switch doItAgain
            case 1 % 繼續 Loop
                continue
            case 0 % 離開 while(doitagain) 紀錄完 換下一張圖
                isRecord = 1;
                break
            case -1 % 系統停止
                return
        end 
    end
    %% 紀錄該值 PPI
    while (isRecord == 1)
        close all
        prompt = "Record PPI. Please Enter Group and Element in form (GXEY): ";
        txt = input(prompt,"s");
        GPattern = caseInsensitivePattern("G");
        EPattern = caseInsensitivePattern("E");
        GString = extractBetween(txt,GPattern,EPattern);
        EString = extractAfter(txt,EPattern);
        GValue = str2double(GString); 
        if isnan(GValue);beep;warning("cannot extract G value. Please try typing again.");continue;end
        EValue = str2double(EString); 
        if isnan(EValue);beep;warning("cannot extract E value. Please try typing again.");continue;end
        LP=[0.25	0.5	    1	    2	    4	    8	    16	    32	    64	    128	    256	    512
        0.281	0.561	1.12	2.24	4.49	8.98	17.96	35.9	71.8	143.7	287.4	574.7
        0.315	0.63	1.26	2.52	5.04	10.08	20.16	40.3	80.6	161.3	322.5	645.1
        0.354	0.707	1.41	2.83	5.66	11.31	22.63	45.3	90.5	181	    362	    724.1
        0.397	0.794	1.59	3.17	6.35	12.7	25.4	50.8	101.6	203.2	406.4	812.7
        0.445	0.891	1.78	3.56	7.13	14.25	28.51	57	    114	    228.1	456.1	912.3]; % unit: linepair
        LP2PPI = 25.3996./(1./(LP*2)); % 轉成PPI
        FinalPPI = LP2PPI(EValue,GValue+3);
        disp(strcat("極限解析度: ",num2str(FinalPPI)," PPI"))
        cprintf('key',"Press any key to continue...\n")
        pause
        isRecord = 0;
    end
    
end
disp("程式結束")
%% function
function doItAgain = AskMeQuit()
    answer = questdlg("Do it again?",'','Yes','No','Quit','No');
    switch answer
        case 'Yes'
            doItAgain = 1;
        case 'No'
            doItAgain = 0;
        case 'Quit'
            doItAgain = -1;
            disp("系統停止")
        otherwise
            doItAgain = -1;
            disp("系統停止")
    end
end