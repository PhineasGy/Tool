% 世界座標轉到面板座標
function pointAtPanel_set = World2Panel(pointAtWorld_set,pixelSize,padPanelLengthVer,padPanelLengthHor)
    % 2023 Emeth General Method
    point_num = size(pointAtWorld_set,2);
    pointAtPanel_set = nan(2,point_num);
    for ii = 1:point_num
        pointAtWorld = pointAtWorld_set(:,ii);
        newX = pointAtWorld(1) - (-padPanelLengthVer*0.5+pixelSize*0.5);
        newY = pointAtWorld(2) - (-padPanelLengthHor*0.5+pixelSize*0.5);
        pointAtPixel = ([newX;newY]/pixelSize) + 1;
        pointAtPanel_set(:,ii) = pointAtPixel;
    end
end