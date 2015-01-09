function posStructure = getPosParameters(rawPos)
% Different methods for positioning objects in a slide:
% =====================================================
%
% The whole slide can be divided to at most 6 (10) tiles, each of them is a
% valid 'pos' parameter.
% ------
% NW - NorthWest
% MW - MiddleWest
% SW - SouthWest
% NE - NorthEast
% ME - MiddleEast
% SE - SouthEast
%
% MEN - MittleEastNorth = Is the upper half of MiddleEast
% MWN - MittleWestNorth = Is the upper half of MiddleWest
% MES - MittleEastSouth = Is the lower half of MiddleEast
% MWS - MittleWestSouth = Is the lower half of MiddleWest
%
% A quasi combined case, each of them is a valid 'pos' parameter:
% ------
% NEH - NorthEastHalf = NorthEast + 0.5 (MiddleEast)
% NWH - NorthWestHalf = NorthWest + 0.5 (MiddleWest)
% SEH - SouthEastHalf = SouthEast + 0.5 (MiddleEast)
% SWH - SouthWestHalf = SouthWest + 0.5 (MiddleWest)
%
% It is possible to combine adjacent tiles - they have to form a square to
% work proper, each of them is a valid 'pos' parameter. Combinations are:
% ------
% N  - North     = NorthWest  + NorthEast
% M  - Middle    = MiddleWest + MiddleEast
% S  - Middle    = SouthWest  + SouthEast
% NH - NorthHalf = NEH + NWH
% SH - SouthHalf = SEH + SWH
% W  - West   = NW + MW + SW
% E  - East   = NE + ME + SE
% C - Centred - eq. all tiles together - object will be centred in tile
% C = N + M + S

% All other combinations can be done by e.g. MyCombination = 'N + M'


% myPresentationWidth           = rawPos.myPresentationWidth ;
% myPresentationHeight          = rawPos.myPresentationHeight;
% objectHeight                  = rawPos.objectHeight; 
% objectWidth                   = rawPos.objectWidth;
% height                        = rawPos.userHeight;
% width                         = rawPos.userWidth;
% left                          = rawPos.userLeft;
% top                           = rawPos.userTop;
% defaultWidthPercentage        = rawPos.defaultwidthPercentage;
% defaultWidthDivideTile        = rawPos.defaultWidthDivideTile;   % How to divide West and East tiles. e.g. defaultWidthDivideTile = 40 means 40% West => tiles 60% east-tiles
% defaultOuterGapTileN          = rawPos.defaultOuterGapTileN;     % Default outer gap in pixel for tiles - North
% defaultOuterGapTileS          = rawPos.defaultOuterGapTileS;     % Default outer gap in pixel for tiles - South
% defaultOuterGapTileWE         = rawPos.defaultOuterGapTileWE;    % Default outer gap in pixel for tiles - West and East
% 
% dT = defaultWidthDivideTile;
% gN = defaultOuterGapTileN;
% gS = defaultOuterGapTileS;
% gWE = defaultOuterGapTileWE;
%
%
% Different methods for fitting objects in a slide:
% =================================================
%
%fitWidth  - Fit object completly to Width of tile and keep aspect ratio 
%fitHeight - Fit object completly to Height of tile and keep aspect ratio 
%fitTotal  - Fit object completly to Width AND Height of tile don't keep aspect ratio  
%fitHeightPercentage -  Fit object to for example 80% in Height of the tile and keep aspect ratio 
%fitWidthPercentage (standard) - Fit object to for example 80% in Width of the tile and keep aspect ratio 
%
%


%% We have to get the actual wanted tile (or combination of tiles)
% In addition we have to get fit- Parameter

% Calculate tiledimension and offsets for coordinatetransformation afterwards

%myTileParameters = getTileParameters(myPresentationWidth,myPresentationHeight,stringPos);

myTileParameters = getTileParameters(rawPos);
posStructure     = getObjectInTileParameters(rawPos,myTileParameters);





end



function posStructure = getObjectInTileParameters(rawPos,myTileParameters)

    objectHeight                  = rawPos.objectHeight; 
    objectWidth                   = rawPos.objectWidth;
    height                        = rawPos.userHeight;
    width                         = rawPos.userWidth;
    left                          = rawPos.userLeft;
    top                           = rawPos.userTop;
    
    
    defaultWidthPercentage        = rawPos.defaultwidthPercentage;
    widthPercentageByUser         = rawPos.widthPercentageByUser;
    defaultHeightPercentage       = rawPos.defaultheightPercentage;
    heightPercentageByUser        = rawPos.heightPercentageByUser;
    
    
    leftTile   = myTileParameters(1);
    topTile    = myTileParameters(2);
    widthTile  = myTileParameters(3);
    heightTile = myTileParameters(4);
    
    
    %% First lets see which mode of fitting we want to use
        %Width                - Fit object to certain Width in the tile and keep aspect ratio 
        %Height               - Fit object to certain Height in the tile and keep aspect ratio 
        %Width Height         - Fit object to certain Height and width in the tile - must not keep aspect ratio  
        %Width%               - Fit object to for example 60% in Width of the tile and keep aspect ratio
        %Height%(standard)    - Fit object to for example 60% in Height of the tile and keep aspect ratio
        
%      fittingMode        = rawPos.fittingMode;
%      
%      switch fittingMode
%          case 'fitWidth' % this is basically the case
%              
%          case 'fitHeight'
%          case 'fitTotal'
%          case 'fitHeight%'
%          case 'fitWidth%'
%      end
     
     
    %% Calculate %[top, left, width and height] of object in tile coordinate system
    

    
    if isempty(width)
        
        if widthPercentageByUser ~= 0
            
            %%The user setted a Width%
            width = round(widthTile*defaultWidthPercentage/100);
            
        end
        
        
    end
    

    if isempty(height) %%No user input
        
        if heightPercentageByUser == 0
            
            if isempty(width) && widthPercentageByUser == 0 %No width was setted by the user at all - no absolute no percentage => so we calculate it
               height       = round(heightTile*defaultHeightPercentage/100);
               aspectratio  = objectWidth/objectHeight;
               width        = round(height*aspectratio);
            else
                height = round(heightTile*defaultHeightPercentage/100);
            end
            
            if widthPercentageByUser ~=0
                %The user wants to set the Width%
                width        = round(widthTile*defaultWidthPercentage/100);
                aspectratio  = objectHeight/objectWidth;
                height       = round(width*aspectratio);
            end
            
        else
            %%The user setted a Height%
            height = round(heightTile*defaultHeightPercentage/100);
            
            if isempty(width)
               aspectratio  = objectWidth/objectHeight;
               width        = round(height*aspectratio);
            end
        end
        
    end
    
    
    
% %     if isempty(width)
% %     %We set the width by default to 60% of the tile
% %     width = round(widthTile*defaultWidthPercentage/100);
% %     
% %     else
% %         %Do nothing - keep the value from user
% %     end
% % 
% %     if isempty(height)
% %         %Calculate heigth from width to keep the right aspectratio
% %         aspectratio  = objectHeight/objectWidth;
% %         height       = round(width*aspectratio);
% %     else
% %         %Do nothing - keep the value from user
% %     end

    if isempty(top)
        %Calculate horizontal centering
        top = round((heightTile-height)/2);
    else
        %Do nothing - keep the value from user
    end

    if isempty(left)
        %Calculate vertical centering
        left = round((widthTile-width)/2);

    else
        %Do nothing - keep the value from user
    end
    
    
    %% Now lets transform back to the coordinate system of the slide - it's easy ;-)
    
    posStructure.top    = top + topTile;
    posStructure.left   = left + leftTile;
    posStructure.width  = width;
    posStructure.height = height;

end




function myTileParameters = getTileParameters(rawPos)

%Get data

stringPos = rawPos.stringPos ;


%First we have to split the pos String into all subtiles
    try

    pos = stringSplit(stringPos,'+',1);

    catch myError
        % Wrong format
        display(['Wrong Format']);
    end
    
    %Translate all inputs to subtiles
    for ii=1:length(pos)
        switch pos{ii}
            case 'N'
                pos{ii} = 'NW+NE';
            case 'E'
                pos{ii} = 'NE+ME+SE';
            case 'S'
                pos{ii} = 'SW+SE';
            case 'W'
                pos{ii} = 'NW+MW+SW';
            case 'M'
                pos{ii} = 'MW+ME';
            case 'C'
                pos{ii} = 'NW+NE+MW+ME+SW+SE';
            case 'NH'
                pos{ii} = 'NW+NE+MEN+MWN';
            case 'SH'
                pos{ii} = 'SW+SE+MES+MWS';  
                
            case 'NEH'
                pos{ii} = 'NE+MEN';
            case 'SEH'
                pos{ii} = 'SE+MES';
                
            case 'NWH'
                pos{ii} = 'NW+MWN';
            case 'SWH'
                pos{ii} = 'SW+MWS';
        end
    end
    
    % Now we have to split the string in each element again so that each
    % element has its own field
    
    tempCounter = 1;
    for ii=1:length(pos)
        tempStruc = stringSplit(pos{ii},'+',1);
        
        for jj=1:length(tempStruc)
            
            subtiles{tempCounter} = tempStruc{jj};
            tempCounter = tempCounter+1;
            
        end
        
    end
    
    % No we have a structure in which each element defines a subtile of the
    % wanted tile or even the whole presentation
    myTileParameters = translateSubtilesToParameters(rawPos,subtiles);

end


function myTileParameters = translateSubtilesToParameters(rawPos,subtiles)

%Get data

myPresentationWidth           = rawPos.myPresentationWidth ;
myPresentationHeight          = rawPos.myPresentationHeight;


defaultWidthDivideTile        = rawPos.defaultWidthDivideTile;   % How to divide West and East tiles. e.g. defaultWidthDivideTile = 40 means 40% West => tiles 60% east-tiles
defaultOuterGapTileN          = rawPos.defaultOuterGapTileN;     % Default outer gap in pixel for tiles - North
defaultOuterGapTileS          = rawPos.defaultOuterGapTileS;     % Default outer gap in pixel for tiles - South
defaultOuterGapTileWE         = rawPos.defaultOuterGapTileWE;    % Default outer gap in pixel for tiles - West and East

dT = defaultWidthDivideTile;
gN = defaultOuterGapTileN;
gS = defaultOuterGapTileS;
gWE = defaultOuterGapTileWE;

%%%


%Lets create a structure of array that include left,top,width,height for
%each tile in the order:
% NW - NorthWest          1                                    
% MW - MiddleWest         2
% SW - SouthWest          3
% NE - NorthEast          4
% ME - MiddleEast         5
% SE - SouthEast          6

% MEN - MiddleEastNorth   7
% MWN - MiddleWestNorth   8
% MES - MiddleEastSouth   9
% MWS - MiddleWestSouth   10

%

stringSubtileLable = {'NW','MW','SW','NE','ME','SE','MEN','MWN','MES','MWS'};

%% First we shrink the effective slide concerning to the defined gaps

myPresentationWidthEffective  = myPresentationWidth-2*gWE;  %working
myPresentationHeightEffective = myPresentationHeight-gN-gS; %working

%% Now we define tops, lefts, widths and heights in the effective coordinate system of the slide

%Widths
widthNorthWest       =  round(myPresentationWidthEffective*dT/100);
widthMiddleWest      =  round(myPresentationWidthEffective*dT/100);
widthSouthWest       =  round(myPresentationWidthEffective*dT/100);
widthMiddleWestNorth =  round(myPresentationWidthEffective*dT/100);
widthMiddleWestSouth =  round(myPresentationWidthEffective*dT/100);

widthNorthEast       =  round(myPresentationWidthEffective*(1-dT/100));
widthMiddleEast      =  round(myPresentationWidthEffective*(1-dT/100));
widthSouthEast       =  round(myPresentationWidthEffective*(1-dT/100));
widthMiddleEastNorth =  round(myPresentationWidthEffective*(1-dT/100));
widthMiddleEastSouth =  round(myPresentationWidthEffective*(1-dT/100));

%Heights
heightNorthWest      =  round(myPresentationHeightEffective/3);
heightMiddleWest     =  round(myPresentationHeightEffective/3);
heightSouthWest      =  round(myPresentationHeightEffective/3);
heightNorthEast      =  round(myPresentationHeightEffective/3);
heightMiddleEast     =  round(myPresentationHeightEffective/3);
heightSouthEast      =  round(myPresentationHeightEffective/3);

heightMiddleWestNorth = round(myPresentationHeightEffective/6);
heightMiddleWestSouth = round(myPresentationHeightEffective/6);
heightMiddleEastNorth = round(myPresentationHeightEffective/6);
heightMiddleEastSouth = round(myPresentationHeightEffective/6);

%Lefts
leftNorthWest       = 0;
leftMiddleWest      = 0;
leftSouthWest       = 0;
leftMiddleWestNorth = 0;
leftMiddleWestSouth = 0;

leftNorthEast       = widthNorthWest;
leftMiddleEast      = widthMiddleWest;
leftSouthEast       = widthSouthWest;
leftMiddleEastNorth = widthMiddleWestNorth;
leftMiddleEastSouth = widthMiddleEastNorth;

%Tops

topNorthWest        = 0;
topNorthEast        = 0;

topMiddleWest       = heightNorthWest;
topMiddleEast       = heightNorthWest;
topMiddleWestNorth  = heightNorthWest;
topMiddleEastNorth  = heightNorthWest;

topMiddleWestSouth  = heightNorthWest + heightMiddleWestNorth;
topMiddleEastSouth  = heightNorthWest + heightMiddleWestNorth;

topSouthWest        = heightNorthWest + heightMiddleWest;
topSouthEast        = heightNorthWest + heightMiddleWest;


%% Now we translate the structure subtiles into a structure where each elemnt is a array tops, lefts, widths and heights for the subtile declared in subtiles
%stringSubtileLable = {'NW','MW','SW','NE','ME','SE','MEN','MWN','MES','MWS'};
tempCounter = 1;

subtilesCoordinates{1} = [0,0,0,0];

choosenVersion = 1;
heightMatrixOne = zeros(3,2); %Gives us a basic overview of the choosen tiles and their height
heightMatrixTwo = zeros(4,2); %Gives us a basic overview of the choosen tiles and their height
widthMatrixOne = zeros(3,2); %Gives us a basic overview of the choosen tiles and their width
widthMatrixTwo = zeros(4,2); %Gives us a basic overview of the choosen tiles and their width

     % a(1,2) = 1
     %============
     % 0     height
     % 0     0
     % 0     0

for ii=1:length(subtiles)

    switch subtiles{ii}                        %[tops, lefts, widths and heights]
        case 'NW'
            subtilesCoordinates{tempCounter} = [topNorthWest,leftNorthWest,widthNorthWest,heightNorthWest];
            tempCounter = tempCounter + 1;
            heightMatrixOne(1,1) = heightNorthWest;
            heightMatrixTwo(1,1) = heightNorthWest;
            widthMatrixOne(1,1)  = widthNorthWest;
            widthMatrixTwo(1,1)  = widthNorthWest;
        case 'MW'
            subtilesCoordinates{tempCounter} = [topMiddleWest,leftMiddleWest,widthMiddleWest,heightMiddleWest];
            tempCounter = tempCounter + 1;
            heightMatrixOne(2,1) = heightMiddleWest;
            %heightMatrixTwo not possible
            widthMatrixOne(2,1)  = widthMiddleWest;
            %widthMatrixTwo not possible
        case 'SW'
            subtilesCoordinates{tempCounter} = [topSouthWest,leftSouthWest,widthSouthWest,heightSouthWest];
            tempCounter = tempCounter + 1;
            heightMatrixOne(3,1) = heightSouthWest;
            heightMatrixTwo(4,1) = heightSouthWest;
            widthMatrixOne(3,1)  = widthSouthWest;
            widthMatrixTwo(4,1)  = widthSouthWest;
        case 'NE'
            subtilesCoordinates{tempCounter} = [topNorthEast,leftNorthEast,widthNorthEast,heightNorthEast];
            tempCounter = tempCounter + 1;
            heightMatrixOne(1,2) = heightNorthEast;
            heightMatrixTwo(1,2) = heightNorthEast;
            widthMatrixOne(1,2)  = widthNorthEast;
            widthMatrixTwo(1,2)  = widthNorthEast;
        case 'ME'
            subtilesCoordinates{tempCounter} = [topMiddleEast,leftMiddleEast,widthMiddleEast,heightMiddleEast];
            tempCounter = tempCounter + 1;
            heightMatrixOne(2,2) = heightMiddleEast;
            %heightMatrixTwo not possible
            widthMatrixOne(2,2)  = widthMiddleEast;
            %widthMatrixTwo not possible
        case 'SE'
            subtilesCoordinates{tempCounter} = [topSouthEast,leftSouthEast,widthSouthEast,heightSouthEast];
            tempCounter = tempCounter + 1;
            heightMatrixOne(3,2) = heightSouthEast;
            heightMatrixTwo(4,2) = heightSouthEast;
            widthMatrixOne(3,2)  = widthSouthEast;
            widthMatrixTwo(4,2)  = widthSouthEast;
        case 'MEN'
            subtilesCoordinates{tempCounter} = [topMiddleEastNorth,leftMiddleEastNorth,widthMiddleEastNorth,heightMiddleEastNorth];
            tempCounter = tempCounter + 1;
            %heightMatrixOne not possible
            heightMatrixTwo(2,2) = heightMiddleEastNorth;
            %widthMatrixOne not possible
            widthMatrixTwo(2,2)  = widthMiddleEastNorth;
            choosenVersion = 2;
        case 'MWN'
            subtilesCoordinates{tempCounter} = [topMiddleWestNorth,leftMiddleWestNorth,widthMiddleWestNorth,heightMiddleWestNorth];
            tempCounter = tempCounter + 1;
            %heightMatrixOne not possible
            heightMatrixTwo(2,1) = heightMiddleWestNorth;
            %widthMatrixOne not possible
            widthMatrixTwo(2,1)  = widthMiddleWestNorth;
            choosenVersion = 2;
        case 'MES'
            subtilesCoordinates{tempCounter} = [topMiddleEastSouth,leftMiddleEastSouth,widthMiddleEastSouth,heightMiddleEastSouth];
            tempCounter = tempCounter + 1;
            %heightMatrixOne not possible
            heightMatrixTwo(3,2) = heightMiddleEastSouth;
            %widthMatrixOne not possible
            widthMatrixTwo(3,2)  = widthMiddleEastSouth;
            choosenVersion = 2;
        case 'MWS'
            subtilesCoordinates{tempCounter} = [topMiddleWestSouth,leftMiddleWestSouth,widthMiddleWestSouth,heightMiddleWestSouth];
            tempCounter = tempCounter + 1;
            %heightMatrixOne not possible
            heightMatrixTwo(3,1) = heightMiddleWestSouth;
            %widthMatrixOne not possible
            widthMatrixTwo(3,1)  = widthMiddleWestSouth;
            choosenVersion = 2;
    end
end


%% Now lets combine all subtilecoordinates to one tile - we will just take the maximum values of each tile
% we are searching for minimum top, minimum left and minimum top,
% maximum left => Because this will define our corner subtiles

%[tops, lefts, widths and heights]

if choosenVersion == 2
    heightMatrix = heightMatrixTwo;
    widthMatrix  = widthMatrixTwo;
else
    heightMatrix = heightMatrixOne;
    widthMatrix  = widthMatrixOne;
end

leftMax     = -9999;
topMax      = -9999;

leftMin     = 9999;
topMin      = 9999;

indexMin    = 1;
indexMax    = 1;

for ii=1:length(subtilesCoordinates)
    
    if subtilesCoordinates{ii}(2)<= leftMin && subtilesCoordinates{ii}(1)<= topMin
        leftMin = subtilesCoordinates{ii}(2);
        topMin  = subtilesCoordinates{ii}(1);
        indexMin = ii;
    end
    
% %     if subtilesCoordinates{ii}(2)>= leftMax && subtilesCoordinates{ii}(1)>= topMax
% %         leftMax = subtilesCoordinates{ii}(2);
% %         topMax  = subtilesCoordinates{ii}(1);
% %         indexMax = ii;
% %     end
    
end

% If indexMin eq. indexMax there is only one subtile

%% Now calculate the actual parameters and do a coordinate back transformation


leftTile    = subtilesCoordinates{indexMin}(2);
topTile     = subtilesCoordinates{indexMin}(1);


%For getting the total height we add all values in each collum of the
%heighmatrix togther and set the maximum to the tile height

heightTile = max(sum(heightMatrix));

%For getting the total width we add all values in each row of the
%widthmatrix togther and set the maximum to the tile width

widthTile = max(sum(widthMatrix'));


% if indexMin ~= indexMax
%     widthTile   = subtilesCoordinates{indexMin}(3) + subtilesCoordinates{indexMax}(3);
%     %heightTile  = subtilesCoordinates{indexMin}(4) + subtilesCoordinates{indexMax}(4);
% else
%     widthTile   = subtilesCoordinates{indexMin}(3);
%     %heightTile  = subtilesCoordinates{indexMin}(4);
% end

%% Coordinate back transformation

leftTile = leftTile + gWE;
topTile  = topTile  + gN;


myTileParameters = [leftTile, topTile, widthTile, heightTile];

end

%
function splitString = stringSplit(str, deli, deleteDeli)

    if nargin == 1
      deli = ' ';
      deleteDeli = 0;
    end
    
    if nargin == 2
      deleteDeli = 0;
    end
    
    jj = 1;
    lastdeli = 0;
    
    for ii=1:length(str)
        
        if str(ii) == deli
            
            splitString{jj} = str(lastdeli+1:ii-deleteDeli);
            lastdeli = ii;
            jj = jj + 1;
            
        end
        
    end
    
    splitString{jj} = str(lastdeli+1:end);


end