
V = 3840;
H = 2160;

segNum = 4;
directionArray = [1,2,3,4,5]; % M LU LD RU RD

%%
directionStr = ["_(M)","_(LU)","_(LD)","_(RU)","_(RD)"];

for ii = directionArray
    switch ii
        case 1
            locationV = 0;
            locationH = 0;
        case 2
            locationV = -1;
            locationH = -1;
        case 3
            locationV = 1;
            locationH = -1;
        case 4
            locationV = -1;
            locationH = 1;
        case 5
            locationV = 1;
            locationH = 1;
    end
    imcropGY(nan,[100,80],[0+locationV*V/segNum,0+locationH*H/segNum],directionStr(ii))
end