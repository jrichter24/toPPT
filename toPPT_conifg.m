function desiredProperty = toPPT_conifg(identifierProperty)

    %% Default properties for adding figures
    toPPTFigure.defaultSlideNumber       = 'append';
    toPPTFigure.defaultSlideAddMethod    = 'update';
    toPPTFigure.defaultwidthPercentage   = 80; %in percent
    toPPTFigure.widthPercentageByUser    = 0;
    toPPTFigure.defaultheightPercentage  = 100; %in percent
    toPPTFigure.heightPercentageByUser   = 0;
    toPPTFigure.stringPos                = 'C';
    toPPTFigure.defaultWidthDivideTile   = 50; %in percent
    toPPTFigure.defaultOuterGapTileN     = 110; %in px
    toPPTFigure.defaultOuterGapTileS     = 60; %in px
    toPPTFigure.defaultOuterGapTileWE    = 10; %in px
    toPPTFigure.userHeight               = [];
    toPPTFigure.userLeft                 = [];
    toPPTFigure.userTop                  = [];
    toPPTFigure.userWidth                = [];
    toPPTFigure.myPresentationHeight     = [];
    toPPTFigure.myPresentationWidth      = [];
    toPPTFigure.objectHeight             = [];
    toPPTFigure.objectWidth              = [];
    
    
    % Additional for figures in png format
    toPPTFigurePNG                       = toPPTFigure;
    toPPTFigurePNG.defaultMagnify        = 2;
    
    
    
    % Additional for exporting as vector graphics
    toPPTFigureVEC                       = toPPTFigure;
    toPPTFigureVEC.defaultAutoGroupMin   = 30;
    

    
    %% Default values for exporting texts
    toPPTText.isGeneralCommand          = 0; % If we want to save a presentation isGeneralCommand will be set to 1
    
    % Template
    toPPTText.doApplyTemplate           = 0; % By default no template is present
    % Save 
    toPPTText.doSavePPTPath             = 0; 
    toPPTText.doSavePPTFilename         = 0;
    toPPTText.doClose                   = 0;
    
    toPPTText.defaultSlideNumber        = 'current';
    toPPTText.defaultSlideAddMethod     = 'update';
    toPPTText.defaultwidthPercentage    = 100; %in percent
    toPPTText.widthPercentageByUser     = 1;
    toPPTText.defaultheightPercentage   = 100; %in percent
    toPPTText.heightPercentageByUser    = 1;
    toPPTText.stringPos                 = 'W';
    toPPTText.defaultWidthDivideTile    = 50; %in percent
    toPPTText.defaultOuterGapTileN      = 90; %in px
    toPPTText.defaultOuterGapTileS      = 10; %in px
    toPPTText.defaultOuterGapTileWE     = 10; %in px
    toPPTText.defaultTitle              = 'MySlide'; %in percent
    toPPTText.defaultTable              = 'MyTable';   
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
    
    toPPTText.myPresentationHeight      = [];
    toPPTText.myPresentationWidth       = [];
    toPPTText.objectHeight              = [];
    toPPTText.objectWidth               = [];
    
    toPPTText.doSetHyper                = 0;
    toPPTText.defaultHyperlink          = 'MyHyperLink';
    
    %% QR CODE PARTS like source url for jars etc. will be taken from qrcode_config
    toPPTQR.color   = '#000000';
    toPPTQR.BGcolor = '#FFFFFF';
    

    
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