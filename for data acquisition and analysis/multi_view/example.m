%%
clear
clc
close all
%% multi-view coding
% 1. 1D pinhole array (f=G)
% 2. 1 view <=> 1 pixel
% 3. current version: LRA0

%% image
image_mode = 0;
    RDP = 40;

VD = 500;   % VD = VDZ if direct viewing --> for view center Z (mm)
IPD = 60;   % for view size (mm)
p = 0.1;      % lens pitch (mm)
views = 4;  % total view number
n = 1.5;    % lens_n
panelPixelNumberVer = 1080; % pixel number (ver)
panelPixelNumberHor = 1920; % pixel number (hor)
method = "subpixel";

%% pre-processing
view_total_size = IPD * views;
view_each_size = IPD;
view_pitch = view_each_size;
switch method
    case "pixel"
        pixel_size = p / views;
    case "subpixel" % p / views = subpixel size
        pixel_size = 3 * (p / views);
    otherwise
        error("[error] method must be either pixel or subpixel")
end
G = p * VD / (view_total_size);
% effective lens (f=G)
R = (n-1)*G;
H = R - sqrt(R^2-(0.5*p)^2);
panelLengthVer = pixel_size * panelPixelNumberVer;
panelLengthHor = pixel_size * panelPixelNumberHor;

%% Camera Stage
% view center and lens very edge
% lens z = 0
view_array = linspace(-(views-1),views-1,views);
view_center_array = [view_array * IPD/2;linspace(VD,VD,views)];
lens_very_edge_array = [-panelLengthHor/2,+panelLengthHor/2;0,0];
% derive RR and RP center for each lens and each view
RR_z = RDP; RP_z = -G;
RR_center_xcell = nan(2,views);

% View object
V = View(number=views,resolution=[panelPixelNumberVer,panelPixelNumberHor]);
if image_mode == 1
    for vv = 1:views
        for ll = 1:2
            view_center = view_center_array(:,vv);
            lens_center = lens_very_edge_array(:,ll);
            
            % RR region
            RR_center_x = (RDP-view_center(2))*((lens_center(1)-view_center(1))/(lens_center(2)-view_center(2)))+view_center(1);
            RR_center = RR_center_x;
    
            RR_center_xcell(ll,vv) = RR_center;
        end
    end
end

% imaging pad
I = imread("-1Group_25Position.png"); I = imresize(I,[panelPixelNumberVer,panelPixelNumberHor]);
padSizeHor = max([max(0 - RR_center_xcell,[],"all"),max(RR_center_xcell-panelLengthHor,[],"all")]);
padSizeHor = ceil(padSizeHor/pixel_size); if isnan(padSizeHor); padSizeHor = 0;end
padSizeVer = padSizeHor;
padI = padarray(I,[padSizeHor padSizeVer],0,'both');
padPanelPixelNumVer = padSizeVer * 2 + panelPixelNumberVer;
padPanelPixelNumHor = padSizeHor * 2 + panelPixelNumberHor;
padPanelLengthVer = padPanelPixelNumVer*(panelLengthVer/panelPixelNumberVer);
padPanelLengthHor = padPanelPixelNumHor*(panelLengthHor/panelPixelNumberHor);

if image_mode == 1
    for vv = 1:views
        % image
        RR_center_xcell_panel = arrayfun(@(x) round(World2Panel([x;0],pixel_size,padPanelLengthHor,padPanelLengthVer)),RR_center_xcell,'UniformOutput',false);
        RR_ind1 = RR_center_xcell_panel{1,vv}(1);
        RR_ind2 = RR_center_xcell_panel{2,vv}(1);
        % image extract + unpad
        RR_content = padI(padSizeVer+1:end-padSizeVer,RR_ind1:RR_ind2,:);

        RR_content = imresize(RR_content,[panelPixelNumberVer,panelPixelNumberHor/views]);
        V.image{vv} = RR_content;
    end
elseif image_mode == 0
    RR_content = uint8(255*ones(panelPixelNumberVer,panelPixelNumberHor/views,3));
    V.image(:) = {RR_content};
end

%% Arrange Step
V.createII(method=method,only=3)
