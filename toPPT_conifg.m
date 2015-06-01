function desiredProperty = toPPT_conifg(identifierProperty)

    % Before next release (toPPT2.1) try to finish:

    % Documentation
    
    % Another look on hyperlinks
    
    % Todo add export icon to figures
    
    % Comments and (replies no) DONE
    
    % Add video audio objects
    
    % Animations - Transistions between slides
    
    % Start slidshow
    
    % Select language - translate - spelling
    
    % Connect elements with arrows
    
    % Improve postioning of graphs and texts by pixel/percentage position DONE
    
    % Add notes to certain slides
    % Notes: op.NotesMaster.HeadersFooters.Footer.Text = 'Volcano Coffee'
    
    % Bundle slides to sections DONE
    
    % Correct pptfigure for matlab 2015a again :-(
    
    % Make toPPT accept cells with commands. E.g.
    % toPPT('commandChain',{SomeCommands}) DONE
    % The advantage is that no escaping of characters is necessary
    
    % Load presentation DONE
    
    % Start presentation form certain number DONE
    
    % Change format and orientation of presentation DONE
    
    % Password encryption - Protect DONE
    
    % Update qr code quietzone setting
    
    % Tables can use css tags and TEx DONE

    % Tables and normal texxt can use new css tag bg for bacground color DONE



    %% Default properties for adding figures
    toPPTFigure.defaultSlideNumber       = 'append';                        %% User can change value
    toPPTFigure.defaultSlideAddMethod    = 'update';                        %% User can change value
    toPPTFigure.defaultwidthPercentage   = 80; %in percent                  %% User can change value
    toPPTFigure.widthPercentageByUser    = 0;
    toPPTFigure.defaultheightPercentage  = 100; %in percent                 %% User can change value
    toPPTFigure.heightPercentageByUser   = 0;
    toPPTFigure.stringPos                = 'C';                             %% User can change value
    toPPTFigure.defaultWidthDivideTile   = 50; %in percent                  %% User can change value
    toPPTFigure.defaultOuterGapTileN     = 110; %in px                      %% User can change value
    toPPTFigure.defaultOuterGapTileS     = 60; %in px                       %% User can change value
    toPPTFigure.defaultOuterGapTileWE    = 10; %in px                       %% User can change value
    toPPTFigure.userHeight               = [];
    toPPTFigure.userLeft                 = [];
    toPPTFigure.userTop                  = [];
    toPPTFigure.userWidth                = [];
    toPPTFigure.myPresentationHeight     = [];
    toPPTFigure.myPresentationWidth      = [];
    toPPTFigure.objectHeight             = [];
    toPPTFigure.objectWidth              = [];
    toPPTFigure.defaultPosAnker          = 'C';                             %% User can change value
    toPPTFigure.posPercentageByUser      = 0;
    
    % Additional for figures in png format
    toPPTFigurePNG                       = toPPTFigure;
    toPPTFigurePNG.defaultMagnify        = 2;                               %% User can change value
    
    
    
    % Additional for exporting as vector graphics
    toPPTFigureVEC                       = toPPTFigure;
    toPPTFigureVEC.defaultAutoGroupMin   = 30;                              %% User can change value

    
    %% Default values for exporting texts
    toPPTText.isGeneralCommand          = 0; % If we want to save a presentation isGeneralCommand will be set to 1
    
    % Page Setup
    toPPTText.doSetPageFormat           = 0; % By default the default page setup from powerpoint will be used
    toPPTText.doSetPageOrientation      = 0; % By default the default page setup from powerpoint will be used
    
    %toPPTText.pageFormat                = '4:3'; % Available values would be 4:3, 16:9 and so on
    toPPTText.knownPageFormats          = {  '4:3',  '16:9', '16:10', 'Letter','Ledger', 'A3',    'A4',       'B4',      'B5'}; % Title of the format
    toPPTText.knownPageFormatsDesc      = {[10,7.5],[10,5.6],[10,6.2],[8.5,11],[17,11],[14,10.5],[10.8,7.5],[11.8,8.9],[7.8,5.9]} ; % Size of format in the same oder as the title in inch (width,height)
    
    %toPPTText.pageOrientation           = 'landscape'; 
    toPPTText.knownPageOrientations     = {'landscape','portrait','invert'}; % invert means that the current orientation is switched e.g from landscape to portrait or vice versa
    
    %toPPTText.firstSlideNumber          = 1;
    toPPTText.doSetFirstSlideNumber     = 0; % By default we will not overwrite the numbering of the current presentation
    
    % Template
    toPPTText.doApplyTemplate           = 0; % By default no template is present
    % Save 
    toPPTText.doSavePPTPath             = 0; 
    toPPTText.doSavePPTFilename         = 0;
    toPPTText.doClose                   = 0;
    toPPTText.hasSavePassword           = 0;  % By default save presentations without password
    toPPTText.savePassword              = '';
    % Open
    toPPTText.doOpenPPT                 = 0;
    
    % Comment
    toPPTText.addComment                = 0;
    
    toPPTText.defaultSlideNumber        = 'current';                        %% User can change value
    toPPTText.defaultSlideAddMethod     = 'update';                         %% User can change value
    toPPTText.defaultwidthPercentage    = 100; %in percent                  %% User can change value
    toPPTText.widthPercentageByUser     = 1;
    toPPTText.defaultheightPercentage   = 100; %in percent %% User can change value
    toPPTText.heightPercentageByUser    = 1;
    toPPTText.stringPos                 = 'W'; 
    toPPTText.defaultWidthDivideTile    = 50; %in percent                   %% User can change value
    toPPTText.defaultOuterGapTileN      = 90; %in px                        %% User can change value
    toPPTText.defaultOuterGapTileS      = 10; %in px                        %% User can change value
    toPPTText.defaultOuterGapTileWE     = 10; %in px                        %% User can change value
    toPPTText.defaultTitle              = 'MySlide';                        %% User can change value
    toPPTText.defaultTable              = 'MyTable';                        %% User can change value
    toPPTText.doSetTitle                = 0;
    toPPTText.doSetText                 = 0;
    toPPTText.doSetTable                = 0;
    toPPTText.doSetTableS               = 0;
    toPPTText.addBulletPoints           = 1;
    toPPTText.addBulletNumbers          = 0;
    toPPTText.userHeight                = [];
    toPPTText.userLeft                  = [];
    toPPTText.userTop                   = [];
    toPPTText.userWidth                 = [];
    toPPTText.doTexText                 = 1; % By default in each text attribute tex code is searched
    toPPTText.posPercentageByUser       = 0;
    toPPTText.userSection               = [];
    toPPTText.doAddSection              = 0;
    
    toPPTText.myPresentationHeight      = [];
    toPPTText.myPresentationWidth       = [];
    toPPTText.objectHeight              = [];
    toPPTText.objectWidth               = [];
    toPPTText.defaultPosAnker           = 'NW';                             %% User can change value
    
    toPPTText.doSetHyper                = 0;
    toPPTText.defaultHyperlink          = 'MyHyperLink';
    
    %% QR CODE PARTS like source url for jars etc. will be taken from qrcode_config
    toPPTQR.color   = '#000000';                                            %% User can change value
    toPPTQR.BGcolor = '#FFFFFF';                                            %% User can change value
    

    
    %% Now lets return the desired property
    switch identifierProperty
        case 'toPPTFigurePNG'
            desiredProperty = toPPTFigurePNG;
        case 'toPPTFigureVEC'
            desiredProperty = toPPTFigureVEC;
        case 'toPPTText'
            desiredProperty = toPPTText;
        case 'toPPTQR'
            desiredProperty = toPPTQR;
    end
    
end