% toPPT-2.5 by Jens Richter 
% email: jrichter@iph.rwth-aachen.de
%
% The following help describes the parameters in the order they appear in
% example.m file (Please start example.m to see how toPPT performs).
%
% Available inputs for toPPT(input):
% ==================================
%
% Parameter 'setTitle':
% ---------------
%   toPPT('setTitle','Your title') - adds a title to your slide
%
%
% Input figure:
% ---------------
%   toPPT(figure) -  adds a figure with default settings to a new slide.
%           By default figures are added in png-format to a new slide.    
%
%
% Parameter 'pos' and 'pos%':
% ---------------
%   toPPT(figure,'pos','NEH') - adds a figure to NorthEastHalf position of
%           a new slide.
%           => Use 'help getPosParameters' for all available positions.
%
%   Additionally the 'pos' parameter also accepts numerical arrays of two 
%   values where the first value is indicating the x-position in pixel and 
%   the second value is indicating the position in y-position.
%   The parameter 'pos%' ONLY accepts numerical arrays of two 
%   values. Where the first value is indicating a x-position in percentage 
%   (calculated from the current width of the presentation slides) and 
%   the second value is indicating a position in y-position in percentage
%   (calculated from the current height of the presentation slides).
%
%   Example:
%   toPPT(figure,'pos%',[50,70])
%   toPPT(figure,'pos',[250,350])
%
% Parameter 'posAnker' (for use togehter with 'pos' or 'pos%'):
% ---------------
%   The position anker is defining which corner of the figure to be
%   exported by toPPT should match to the specified position (via 'pos' or
%   'pos%'). The default anker point of each figure is its center.
%
%   => Use 'help getPosParameters' for all available positions.
% 
%
% Parameter 'SlideNumber' and 'SlideAddMethod':
% ---------------
%   toPPT(...,'SlideNumber','current') - adds your content to the active slide
%   toPPT(...,'SlideNumber','append') - adds your content to a new slide.
%   toPPT(...,'SlideNumber',3) - adds your content to slide 3. In case 3 is
%           not already existing empty slide will be placed in between.
%   toPPT(...,'SlideNumber','Title of another slide') - updates the slide
%           with 'Title of another slide' with your content. Edit distances will be
%           used so it is not highly sensitive to errors in the spelling of the slide title.
%   toPPT(...,'SlideNumber','Title of another slide','SlideAddMethod','insert')
%           Instead of updating a slide it appends a new slide after the
%           selected slide.
%
%
% Parameter 'existingFigure':
% ---------------
%   Loads an exiting figure. Can be combined with all other commands. The 
%   path can be defined relative or absolute.
%
%   Example:
%   toPPT('existingFigure','C:\myFig.fig');
%
%
% Parameter 'existingImage':
% ---------------
%   Loads an exiting image. Can be combined with all other commands. The 
%   path can be defined relative or absolute.
%
%   Example:
%   toPPT('existingFigure','C:\myFig.png');
%
%
% Input 'Your Text String':
% ---------------
%   toPPT('My text') - adds your text as bullet point.
%           By default text/tables are added to the current slide. 
%
% Input {'Text1','Text2','Text3'}
% ---------------
%   toPPT({'Text1','Text2','Text3'}) - adds each of the elements in the cell as separate bullet point.
%
%
% Parameter 'hasFrame','frameType','frameColor' and 'frameWidth'
% ---------------
%   toPPT('Mytext','hasFrame',1,'frameType','DashDotDot','frameColor','blue','frameWidth',6)
%
%   - Can be used to add an visible frame arround a textBox. For more
%   information on frameTypes check toPPT_conifg under toPPTText.knownTextBoxFramesTitle
%
%
% Parameter 'rotationAngle'
% ---------------
%   toPPT('Mytext','rotationAngle',10)
%
%   - Rotates a textBox by 10 degrees.
%
%
%
% Parameter 'setBulletNumbers'
% ---------------
%   toPPT('My text','setBulletNumbers',1) - uses numbers instead of bullet points.
%
%
% Parameter 'setBullets'
% ---------------
%   toPPT('My text','setBullets',0) - deactivates bullets.
%
%
% Input '<b>Your Text String with css tags </b>':
% ---------------
%   UPDATE: Each cell of a table can accept s-tags and TeX!
%
% 	toPPT('<b>Your Text String with css tags </b>') - it is possible to use
%       css tags to change color, font, size etc. of text.
%       => Use 'help addText' for all available options.
%
%
% Input 'x = \sqrt\x^2 - your TeX code'
% ---------------
%   toPPT('x = \sqrt\x^2- your TeX code') - it is possible to use
%           TeX tags.
%           => Use 'help tex2ppt' for all available options.
%
% Parameter 'TeX':
% ---------------
%   By default in each text that is delivered toPPT will search for
%   TeX-Characters and translates them. Sometimes this behaviour is not
%   desirable. For this purpose the TeX-Interpreter can be turned off
%   temporally.
%
%   Example:
%   toPPT('TeX Interpreter is off for x = \sqrt\x^2','TeX',0);
%
%
% Parameter 'setTable'
% ---------------
%   UPDATE: Each cell can accept s-tags and TeX! 
%   => Use 'help addText' for all available options.
%
%   toPPT('setTable',{stringCellTableCaption, matrix or Vector or cell})
% 
%   Examples -  adding tables with different input:
% 	helpTableCell      = cell(2,3);
%   helpTableCell{1,1} = 'Me is first';
%   helpTableCell{1,2} = 'Me is second';
%   helpTableCell{1,3} = 'Me is third';
%   helpTableCell{2,1} = 'Me is fourth';
%   helpTableCell{2,2} = 'Me is fifth';
%   helpTableCell{2,3} = 'Me is sixth';
% 
%   - It is important to add "helpTableCell" into an extra cell {}
%   toPPT('setTable',{  {'Caption1','Caption2','Caption3'} , {helpTableCell}  } );
%
%   - Auto rotate is a strategy that is automatically used to fit the number of
%   captions to the number of elements.
%   toPPT('setTable',{  {'Caption1','Caption2'},{helpTableCell}  });
% 
%   - It is important to add all vectors into an extra cell {}
%   helperVector1 = [1,2,3];
%   helperVector2 = {'Me is 1','Me is 2','Me is 3'};
% 
%   toPPT('setTable',{  {'Col 1','Col 2'},{helperVector1,helperVector2}  });
% 
%   - With autorotation again:
%   toPPT('setTable',{  {'Col 1','Col 2','Col 3'},{helperVector1,helperVector2}  });
%
% Parameter 'pub':
% ---------------
%   toPPT(...,'pub',0) - deactivates publishing. This option is useful because
%           sometimes depending on the script the toPPT output takes some time
%           and the user does not want to use toPPT while debugging his/her
%           script. 'pub',1 is the default.
%
%
% Parameter 'Height', 'Height%','Width', 'Width%':
% ---------------
%   The parameters Width and Height are referring to absolute dimensions in pixels.
%   The parameters Width% and Height% are referring to relative dimensions in
%   percentage of the tile the figure will be placed in. If the tile is
%   bigger the image will be bigger to in contrast of placing a figure in a
%   smaller tile with the same Width% and Height% settings. If only one of
%   the parameters Width(%) or Height(%) is set the aspect ratio of the
%   figure is persisted.
%
%   Examples:
%   toPPT(figure,'pos','M','Height',200,'Width',100) -  adds a figure to the
%           middle tile. The figure will have a size of 100 px x 200 px.
%   toPPT(figure1,'pos','C','Height%',100) - - adds a figure to the
%           center tile. The figure will have a heigth of 100% of the tile. The
%           width is calculated depending on the aspect ratio.
% 
%
% Parameter 'gapN', 'gapWE' and 'gapS':
% ---------------
%   The parameters gapN (North), gapWE (West) and gapS (South) will define the 
%   gap between the outer slide etch and the start of the active placing area 
%   in pixels. E.g. for not putting a figure overlaying the title of a slide
%   it is important to have a certain padding or gap. 
%
%   Example:
%   toPPT(figure,'gapN',200,'gapWE',150) - will set the gap in the north to
%           200 pixels and to 150 pixels in the West.
%
%
% Parameter 'divide':
% ---------------
%   The user is also able to overwrite the "divide" value for East and West.
%   This means e.g. if "divide" is set to 40 (from 100) all West tiles will 
%   have a width of 40% of the slide and all East tiles
%   will have 60% of the width of the slide. (Gap settings excluded).
%
%   Example:
%   toPPT(figure,'pos','E','Width%',100,'divide',30) - adds a figure to the
%   east tile with 100% width. The divide property is overwritten 
%
%
% Parameter 'm':
% ---------------
%   The parameter 'm' defines the resolution that is used for adding figures
%   as images (png) to a slide. The default value is 2. The higher the
%   value the higher the resolution and the bigger the image is in terms of 
%   disk space.
%   
%   Example:
%   toPPT(figure,'m',1) - adds the image with low resolution  to a new
%       slide.
%
%
% Parameter 'format':
% ---------------
%   By default figures are added as png to a new slide. Sometimes it is
%   desirable to export a figure as a vector image.  
%
%   Example:
%   toPPT(figure,'format','vec') - adds a figure to a new slide in vector
%   format.
%
%
%
% Parameter 'exportMode' and 'exportFormatType':
% ---------------
%   By default figures are added as png by using exportFig. The quality of
%   the export has a high level using exportFig. In case of generating
%   presentations with a lot of images the export is quite time consuming and 
%   the resulting presentation has a huge file size.
%   By setting exportMode to "matlab" instead of "exportFig" (default) the
%   internal matlab functions will be used for exporting the figure what is
%   usually much faster and the resulting presentation is much smaller.
%   When using "matlab" for exporting an exportFormatType can be specified 
%   what will define the export type.
%
%   The following exportFormatTypes are possible:
%   'jpeg','png','tiff','tiffn','meta','bmpmono','bmp','bmp16m','bmp256',
%   'hdf','pbm','pbmraw','pcxmono','pcx24b','pcx256','pcx16','pgm',
%   'pgmraw','ppm' and 'ppmraw'.
%
%   'meta' is the default exportFormatType.
%
%
%   Example:
%   toPPT(figure,'exportMode','matlab');
%   toPPT(figure,'exportMode','matlab','exportFormatType','bmp');
%   
%
% Parameter 'addSection':
% ---------------
%   This parameter is defined for adding sections to the active
%   presentation. For adding sections 'SlideAddMethod' also accepts 'before'
%   and 'after' as argument.
%
%   Example:
%   toPPT('addSection','My first section','SlideNumber',3,'SlideAddMethod','before');
%   - will add a section called "My first section" in front of slide 3.
%   
%
% Parameter 'asComment':
% ---------------
%   For adding comments to the active slide use:
%
%   Example:
%   toPPT('asComment',{'Your Name','Your initials','Your comment'});
%
%
% Parameter 'setPageOrientation' and 'setPageFormat':
% ---------------
%   The parameter setPageOrientation is used to set the presentation to
%   'landscape', 'portrait' or 'invert' (this is just inverting the current
%   orientation).
%
%   Example:
%   toPPT('setPageOrientation','landscape');
%
%   The parameter setPageFormat is used to define the desired format e.g.
%   16:9. Possible values are '4:3','16:9','16:10','Letter','Ledger',
%   'A3','A4','B4' and 'B5'.
%
%   Example:
%   toPPT('setPageFormat','16:10');
%
%   Annotation: It is recommended to use 'setPageOrientation' 
%   and 'setPageFormat' before adding any content to your presentation.
%
% Parameter 'commandChain':
% ---------------
%   A commandChain can be executed as a chain of commands within an cell 
%   array. CommandChains can be used to design a presentation e.g. within in other
%   script file. The advantage is that no escaping is necessary and the
%   presentation can be easily bundled.
%
%   Example:
%   commandsCell{1} = {figure1,'Width%',50,'pos','ME'};
%   commandsCell{2} = {{'Text1','Text2','Text3'}};
%   commandsCell{3} = {'setTitle','Example using commandChains'};
% 
%   for ii=1:numel(commandsCell)
%     toPPT('commandChain',commandsCell{ii});
%   end
%
%   Annotation: In practice big CommandChains can exhibit the available
%   physical memory. Especially when a lot of figures are put into the
%   commandChain.
%
%
% Parameters for QR-Code generation:
% ---------------
% Parameter 'QR-Size':
% ---------------
% Syntax:
%   toPPT(message, 'QR-Size',[117,117]);
%   or equally:
%   toPPT(message, 'QR-Size',117);
%
% Description:
%   Generates a QR code of the defined Size. The size has to match the
%   17+4N rule (e.g. 17+4*25 = 117).
%
%
% Parameter 'QR-Version':
% ---------------
% Syntax:
%   toPPT(message, 'QR-Version',10);
%
% Description:
%   Generates a QR code of the defined QR code version.
%
%
% Parameter 'QR-ErrQuality':
% ---------------
% Syntax:
%   toPPT(message, 'QR-ErrQuality','L');
%
% Description:
%   ErrQuality accepts one of the following strings "L" or "M" or "H" or
%   "Q".
% Definition of error correction codes (taken from
% http://en.wikipedia.org/wiki/QR_code):
%   Level L (Low) 	7% of code words can be restored.
%   Level M (Medium) 	15% of code words can be restored.
%   Level Q (Quartile)[40] 	25% of code words can be restored.
%   Level H (High) 	30% of code words can be restored.
%
%
% Parameter 'QR-CharacterSet':
% ---------------
% Syntax:
%   toPPT(message, 'QR-CharacterSet','UTF-8');
%
% Description:
%   Defines the character set that will be used. See qrcode_config.m for a
%   full list of available character sets.
%
%
% Parameter 'QR-DownloadJars':
% ---------------
% Syntax:
%   toPPT(someMessage,'QR-DownloadJars', 1);
%
% Description:
%   This will download the necessary Jar files. So it is possible to use
%   qrcode_gen statically instead of importing the jars from the internet.
%
%
% Parameter 'QR-BgColor':
% ---------------
% Syntax:
%   toPPT(someMessage,'QR-BgColor', 'red');
%   or equally:
%   toPPT(someMessage,'QR-BgColor', '#FF0000');
%
% Description:
%   This will set the background color of the QR code. The input needs to
%   be a string specifing a known color or a HEX code with leading #. 
%
%
% Parameter 'QR-color':
% ---------------
% Syntax:
%   toPPT(someMessage,'QR-color', 'blue');
%   or equally:
%   toPPT(someMessage,'QR-color', '#00FF00');
%
% Description:
%   This will set the color of the QR code. The input needs to
%   be a string specifing a known color or a HEX code with leading #. 
%
%
% Parameter 'applyTemplate':
% ---------------
%   It is possible to apply a template (e.g. potx format) to the current
%   presentation. It is necessary to use an absolute path for the template
%   file.
%
%   Example:
%   templatePath = [pwd,'\IPH Slides_EN.potx'] - The template path has to be absolute!
%   toPPT('applyTemplate',templatePath)
%
%
% Parameter 'openExisting':
% ---------------'',''C:\myPresentation.pptx''
%   In case it is desired to add content to an already existing
%   presentation this parameter can be used to load a presentation.
%
%   Example:
%   toPPT('openExisting','C:\myDeeplyDesiredPresentation.pptx');
%
%   Annotation: Password protected presentation can NOT be loaded.
%
% Parameter 'savePath' and 'saveFilename':
% ---------------
%   There are three ways of saving
%
%   Case 1: We only define the 'savePath' as argument:
%     => Presentation will be saved in path as 'The Current Date and Time
%     _new_presentation.pptx'.
%
%   Case 2: We only define the saveFilename as argument:
%     => Presentation will be saved in 'pwd' with the desired filename.
%
%   Case 2: We define savePath and saveFilename as argument:
%     => Presentation will be saved in savePath with the desired filename.
%
%   Example:
%   savePath = pwd;
%   saveFilename = 'My first presentation with toPPT';
%   toPPT('savePath',savePath,'saveFilename',saveFilename) - saves the
%   current presentation.
%
%
% Parameter 'savePassword':
% ---------------
%   This parameter can be used for defining a password for the active
%   presentation to be saved. This parameter can only be used together with
%   the parameters 'savePath' and/or 'saveFilename':
%
%   Example:
%   savePath = pwd;
%   saveFilename = 'My first presentation with toPPT that is protected';
%   toPPT('savePath',savePath,'saveFilename',saveFilename,'savePassword','mySecret') - saves the
%   current presentation.
%
%
% Parameter 'close':
% ---------------
%   The parameter 'close' can be used to close the current presentation. The
%   presentation is NOT automatically saved. This option is useful if it
%   necessary to create multiply presentation instead of a single
%   presentation.
%
%   Example:
%   toPPT('close',1) - closes the current presentation


function toPPT(varargin)


%% toPPT-beta2 by Jens Richter 
% email: jrichter@iph.rwth-aachen.de

%% This little script can be seen as an addon to existing scripts 
% that tries to execute them in a stable reproducable way. In addition it has its own
% "placing" algorithm and a basic "addText" ability (for Tables, Texts, etc.)

%% CREDIT LIST - Acknowledgment
%% External scripts:
%1) export_fig-script by Oliver Woodford
    % http://www.mathworks.com/matlabcentral/fileexchange/23629-exportfig
    % is used to generate pictures (png) from the desired figures
    
%2) pptfigure.m by Dmitriy Aronov
    % http://www.mathworks.com/matlabcentral/fileexchange/30124-smart-powerpoint-exporter/content/pptfigure.m
    % is used for genereating editable figures and an extracted part of it is used as tex2ppt converter
    
%3) RGB triple of color name, version 2 by Kristjan Jonasson
    % http://www.mathworks.com/matlabcentral/fileexchange/24497-rgb-triple-of-color-name--version-2
    % is used to transfrom colornames to hex values etc.
    
%4) Edit Distance Algorithm by  Reza Ahmadzadeh
    % http://www.mathworks.com/matlabcentral/fileexchange/39049-edit-distance-algorithm
    % is used to calculate the edit distance for placing a new slide near
    % an other slide with a certain title.

    
 
%% toPPT Start
    
%% We have to get the values that the user setted - to either forward them to the used external scripts or to process them

%% First we see if it is necessary to do someting at all
doPublishing = 1; %default is yes = 1


%% Do we want to work on a command  chain? Basically toPPT()

isCommandChain = 0;
hasNoErrors    = 1;

structMyArgument = deleteFromArgumentAndGetValue(varargin,'commandChain');
if ~isempty(structMyArgument.value)
    commandChain     = structMyArgument.value;
    varargin         = structMyArgument.arguments;
    isCommandChain   = 1;
end
    

if isCommandChain
    % The argument (or second parameter of toPPT has all commands as cell => lets check that)
    if iscell(commandChain)
        % Only one argument and the argument is a cell itself
        
        % Overwrite argument varagin with command chain 
        varargin = commandChain;

    else
        hasNoErrors = 0;
        warning('The command chain (COC) needs to be defined as cell. E.g COC{1} = "setTitle",  COC{2} = "This is my test title" or COC = {"setTitle","This is my test title"}... ')
    end
end



%%% pub %%% e.g. NE
structMyArgument = deleteFromArgumentAndGetValue(varargin,'pub');
if ~isempty(structMyArgument.value)
    doPublishing     = structMyArgument.value;
    varargin         = structMyArgument.arguments;
end
%%%




%%%

if doPublishing && hasNoErrors %uncomment later
    

    
    
%% In this section we will have to distinguish between text and figures

%%% Check if we want to translate text to QR

textAsQRCode = 0; % By default we do NOT want to translate text in QR-Code

structMyArgument         = deleteFromArgumentAndGetValue(varargin,'TextAsQR');
if ~isempty(structMyArgument.value)
    textAsQRCode        = structMyArgument.value;
    varargin            = structMyArgument.arguments;
end

asExistingFigure = 0; % By default we do not want to load an existing figure

structMyArgument         = deleteFromArgumentAndGetValue(varargin,'existingFigure');
if ~isempty(structMyArgument.value)
    figurePath2Load     = structMyArgument.value;
    asExistingFigure    = 1;
    varargin            = structMyArgument.arguments;
end

asExistingImage = 0; % By default we do not want to load an existing figure

structMyArgument         = deleteFromArgumentAndGetValue(varargin,'existingImage');
if ~isempty(structMyArgument.value)
    imagePath2Load      = structMyArgument.value;
    asExistingImage     = 1;
    varargin            = structMyArgument.arguments;
end

if textAsQRCode
    %if strcmp(version('-release'),'2014b') || strcmp(version('-release'),'2014a')
    if matlabVersionChecker('2014a')
        % Get additional parameters
        %%% QR code parameters %%% Parameter 'Size':Parameter 'Version':
        %%% 'ErrQuality': 'CharacterSet': 'DownloadJars':
        
        qrPropertySetVector = cell(5,2);
        qrPropertySetVector{1,1} = 'Size';
        qrPropertySetVector{2,1} = 'Version';
        qrPropertySetVector{3,1} = 'ErrQuality';
        qrPropertySetVector{4,1} = 'CharacterSet';
        qrPropertySetVector{5,1} = 'DownloadJars';
        
        qrPropertyVectorLogical = zeros(5,1);
        
        
        % Size of QR-Code
        myArgQR.sizeSet = 0;
        structMyArgument            = deleteFromArgumentAndGetValue(varargin,'QR-Size');
        if ~isempty(structMyArgument.value)
            myArgQR.qrSize          = structMyArgument.value;
            varargin                = structMyArgument.arguments;
            myArgQR.sizeSet = 1;
            qrPropertyVectorLogical(1) = 1;
            qrPropertySetVector{1,2} = myArgQR.qrSize;
        end
        
        % Version of QR-Code
        myArgQR.versionSet = 0;
        structMyArgument        = deleteFromArgumentAndGetValue(varargin,'QR-Version');
        if ~isempty(structMyArgument.value)
            myArgQR.qrVersion     = structMyArgument.value;
            varargin           = structMyArgument.arguments;
            myArgQR.versionSet = 1;
            qrPropertyVectorLogical(2) = 1;
            qrPropertySetVector{2,2} = myArgQR.qrVersion;
        end
        
        % Error correction quality
        myArgQR.errQSet = 0;
        structMyArgument          = deleteFromArgumentAndGetValue(varargin,'QR-ErrQuality');
        if ~isempty(structMyArgument.value)
            myArgQR.qrErrQuality  = structMyArgument.value;
            varargin              = structMyArgument.arguments;
            myArgQR.errQSet = 1;
            qrPropertyVectorLogical(3) = 1;
            qrPropertySetVector{3,2} = myArgQR.qrErrQuality;
        end
        
        % Character set for QR code
        myArgQR.charSet = 0;
        structMyArgument          = deleteFromArgumentAndGetValue(varargin,'QR-CharacterSet');
        if ~isempty(structMyArgument.value)
            myArgQR.qrCharaterSet = structMyArgument.value;
            varargin              = structMyArgument.arguments;
            myArgQR.charSet = 1;
            qrPropertyVectorLogical(4) = 1;
            qrPropertySetVector{4,2} = myArgQR.qrCharaterSet;
        end
        
        % Download Jars
        myArgQR.downJarSet     = 0;
        structMyArgument           = deleteFromArgumentAndGetValue(varargin,'QR-DownloadJars');
        if ~isempty(structMyArgument.value)
            myArgQR.qrDownloadJars = structMyArgument.value;
            varargin               = structMyArgument.arguments;
            myArgQR.downJarSet     = 1;
            qrPropertyVectorLogical(5) = 1;
            qrPropertySetVector{5,2} = myArgQR.qrDownloadJars;
        end
        
        % Gernerate argument list for qrCode
        qrCodeArgumentCell = cell(sum(qrPropertyVectorLogical)*2,1);
        
        jj = 1;
        
        for ii=1:5
            if(qrPropertyVectorLogical(ii))
                qrCodeArgumentCell{jj} = qrPropertySetVector{ii,1};
                qrCodeArgumentCell{jj+1} = qrPropertySetVector{ii,2};
                jj = jj + 2;
            end
        end
        
        
        qr = qrcode_gen(varargin{1},qrCodeArgumentCell{1:numel(qrCodeArgumentCell)});
        
        % Get default properties for plotting
        toPPTQR = toPPT_conifg('toPPTQR');
        
        
        % Get some plotting infos
        structMyArgument           = deleteFromArgumentAndGetValue(varargin,'QR-Color');
        if ~isempty(structMyArgument.value)
            toPPTQR.color          = structMyArgument.value;
            varargin               = structMyArgument.arguments;
        end
        
        structMyArgument           = deleteFromArgumentAndGetValue(varargin,'QR-BgColor');
        if ~isempty(structMyArgument.value)
            toPPTQR.BGcolor        = structMyArgument.value;
            varargin               = structMyArgument.arguments;
        end
        

        try
            myColormap = [rgb(toPPTQR.color,'decVec'); rgb(toPPTQR.BGcolor,'decVec')]/255;
            
            figureQR = figure;
            imagesc(qr);
            colormap(myColormap);
            axis off;
            axis equal;
            
        catch
            
            warning('Something went wrong when plotting the QR-Code. Do you assigned QR-Color and QR-BgColor as a string? Allowed values are hex colors like #FF0000 (with leading #) and known color names.');
            
            toPPTQR = toPPT_conifg('toPPTQR');
            
            myColormap = [rgb(toPPTQR.color,'decVec'); rgb(toPPTQR.BGcolor,'decVec')]/255;
            
            figureQR = figure;
            imagesc(qr);
            colormap(myColormap);
            axis off;
            axis equal;
            
        end
        
        
        
        
        varargin{1} = figureQR;
        
%         myColormap = [0, 0, 0; 0.5,0.5,0.5];
%         qr = qrcode_gen(message,'Size',97); % you can also specify [97,97] instead
% 
%         figureQR = figure;
%         imagesc(qr);
%         colormap(myColormap);
%         axis off;
%         axis equal;
        
    else
        warning('You need at least version 2014a of matlab to use the QR-code feature in toPPT');
        varargin{1} = 'You need at least version 2014a of Matlab to use the QR-code feature in toPPT.';
    end
    
end

if asExistingFigure
    
    try
        % Load figure
        loadedfig = openfig(figurePath2Load);
        varargin{1} = loadedfig;
    catch
        
        % Check if figure is existing
        isExistingFigure = isequal(exist(figurePath2Load),2);
        
        if isExistingFigure
            warning(['An error occured loading: ',figurePath2Load,'. The file is existing but can not be opened?!']);
        else
            warning(['An error occured loading: ',figurePath2Load,'. The file is not existing or not readable -  Right path? An active presentation will be used instead. If there is no active presentation a new one will be created.']);
        end
    end
    
end


if asExistingImage
    
    try
        % Load figure
        c = imread(imagePath2Load);
        
        figureExistingImage = figure;
        imagesc(c);
        axis off;
        axis equal;
        varargin{1} = figureExistingImage;
        
    catch
        
        % Check if figure is existing
        isExistingFigure = isequal(exist(imagePath2Load),2);
        
        if isExistingFigure
            warning(['An error occured loading: ',imagePath2Load,'. The file is existing but can not be opened?!']);
        else
            warning(['An error occured loading: ',imagePath2Load,'. The file is not existing or not readable -  Right path? An active presentation will be used instead. If there is no active presentation a new one will be created.']);
        end
    end
    
end



pptOutputVersion    = 2; % Output 1 = editable graphic , 2 = png image (default), 3 = text
booleanValidFigure  = 0; % Valid figure, default is no => we will check
booleanValidText    = 0; % Valid text, default is no => we will check
error = 0;

if isempty(varargin) % No argument :-(
    %%error('Sorry you have to give me at least one argument - figure or string/cell of strings');
    % Call the help
    help toPPT;
    return;
elseif ~isempty(varargin)
    %Check first if argument is pure string or cell array of strings
    if ischar(varargin{1}) % Just a single string?
        pptOutputVersion = 3;
        booleanValidText = 1;
    elseif iscell(varargin{1}) % cell?
        if ischar(varargin{1}{1})% string?
            % First element in the cell is a string?
            % => We are not checking the other fields => will lead to an error but faster
            pptOutputVersion = 3;
            booleanValidText = 1;
        elseif iscell(varargin{1}{1}) %% We have cells within cells => is the first object of this cell a string?
            % We want to transform this to a cell with strings
            newStringCell = cell(1,numel(varargin{1}));
            
            for ii=1:numel(varargin{1})
                tempString = '';
                if ischar(varargin{1}{ii})
                    
                    tempString = varargin{1}{ii};
                    
                elseif iscell(varargin{1}{ii})
                    
                    
                    for jj=1:numel(varargin{1}{ii})
                        if ischar(varargin{1}{ii}{jj})
                            tempString = [tempString,varargin{1}{ii}{jj},' '];
                        else
                            % Error - not supported
                            error = 1;
                        end
                    end
                    
                    
                end
                newStringCell{ii} = tempString;
            end
            if error == 0
                booleanValidText = 1;
                pptOutputVersion = 3;
                varargin{1} = newStringCell;
            end
            
        end
        
    else %%Probably a figure
        try % We use try => if an error happens we just catch it
            booleanValidFigure = strcmpi(get(varargin{1},'type'),'figure'); % Will return 1 if true, 0 if false
        catch err
            throwAsCaller(errorGetCause(err));
        end
        if booleanValidFigure % If true lets save the figure to a new variable to distinguish between arguments texts and figure
            myFigure = varargin{1};
        end
    end
    
end


%% Now we want to get all the arguments restructure them for sending them later to the subscripts - First for the case figure

if booleanValidFigure %If we have a valid Figure  go on....
    defaultOutputFormat = 'png'; %% by default we want to export a png - 'vec' will output 
    
    %%% defaultPutputFormat %%% e.g. 'format','vec'
    structMyArgument = deleteFromArgumentAndGetValue(varargin,'format');
    if ~isempty(structMyArgument.value)
        defaultOutputFormat     = structMyArgument.value;
        varargin                = structMyArgument.arguments;
    end
    %%%
    
    switch defaultOutputFormat
        case 'png'
            pptOutputVersion = 2; % is the default case for figures
        case 'vec'
            pptOutputVersion = 1;
    end
    
    % Now we want to restructure the remaining arguments - if there are
    % some length(varargin)>1
    
    if length(varargin)>1
        
        vararginStructure = cell(length(varargin)-1);
        
        for ii=2:length(varargin) % We start at ii=2 to leave out the figure
           vararginStructure{ii-1} = varargin{ii};
        end
        
    else %Outerwise lets just send a ''
        vararginStructure = cell(1);
        vararginStructure{1} = '';
    end
    
end


%% Now we want to get all the arguments and restructure them for sending them later to the subscripts - Second for the case text
if booleanValidText %If we have a valid Text  go on....

   if ~isempty(varargin)
       
       vararginStructure = cell(length(varargin));
        
       for ii=1:length(varargin) % We start at ii=1
          vararginStructure{ii} = varargin{ii};
       end
        
   else %Outerwise lets just send a ''
        vararginStructure = cell(1);
        vararginStructure{1} = '';
   end
    
end

%% Now we are going to distinguish between arguments that belongs to this script or to the subscripts


%%Default values - myArgumentsImage. (postioning, etc)

if pptOutputVersion == 2
    
    % only for png
    
    myArgumentsImagePNG = toPPT_conifg('toPPTFigurePNG');
    
    %%Overwrite values (if existing) with user input
    myArgumentsImagePNG.pptOutputVersion = pptOutputVersion;
    
    myArgumentsImagePNG = getValuesFromArgument(myArgumentsImagePNG,vararginStructure,'png');
    

end

if pptOutputVersion == 1

    %only for vector

    myArgumentsImageVEC = toPPT_conifg('toPPTFigureVEC');
    
    myArgumentsImageVEC.pptOutputVersion = pptOutputVersion;
    
    %%Overwrite values (if existing) with user input
    myArgumentsImageVEC = getValuesFromArgument(myArgumentsImageVEC,vararginStructure,'vec');
end

if pptOutputVersion == 3
    
    myArgumentsText = toPPT_conifg('toPPTText');
    
    myArgumentsText.pptOutputVersion = pptOutputVersion;

    %%Overwrite values (if existing) with user input
    myArgumentsText = getValuesFromArgument(myArgumentsText,vararginStructure,'text');
    
end

 
%% For better user experience we will check the user input in a "try and error and change" manner => Try user input => Error => Change Input => Hopefully no error

switch pptOutputVersion
    case 1 % Vector output
        try
            %addVectorGraphic(myFigure,varginStructure);
            addVectorGraphic(myFigure,myArgumentsImageVEC);
        catch err
            warning('Your figure can not be added as editable image - We will try to export it as png');
            myArgumentsImageVEC.defaultMagnify = 2;
            
            try
                addBitmapGraphic(myFigure,myArgumentsImageVEC);
                display('Export as png : Sucessfull');
                
            catch myError
                error('Try to export as png : Not sucessfull');
            end
        end
    case 2
        try
            addBitmapGraphic(myFigure,myArgumentsImagePNG);
        catch myError
            display('Your figure can not be added as png image');
        end
    case 3 %Text
        try
             % to addText also gerneal behaivour is adddresed like saving
             % etc.
             
             addText(myArgumentsText);
             
             
        catch myError
            display('Your command was not performed succefully!');
        end
end



end

end



function exceptionCaused = errorGetCause(exception)


switch exception.identifier
    case 'MATLAB:class:InvalidHandle'
        %error('toPPT:invalidHandle','Possibly the figure you are trying to add is already deleted or closed!');
        causeException = MException(exception.identifier, 'Possibly the figure you are trying to add is already deleted or closed!');
        exceptionCaused = addCause(exception,causeException);
        %throwAsCaller(exceptionCaused);
        
    otherwise
        causeException = MException(exception.identifier, 'Unknown error');
        exceptionCaused = addCause(exception,causeException);
end

end


function myArg = getValuesFromArgument(myArg,arguments,out)

    if strcmp(out,'vec') || strcmp(out,'png')
        
        
        %%% exportMode %%%
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'exportMode');
        if ~isempty(structMyArgument.value)
            myArg.defaultExportMode  = structMyArgument.value;
            arguments        = structMyArgument.arguments;
        end
        
        
        %%%toPPTFigure.defaultExportFormatType
        %%% ExportFormatType %%%
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'exportFormatType');
        if ~isempty(structMyArgument.value)
            myArg.defaultExportFormatType  = structMyArgument.value;
            arguments                      = structMyArgument.arguments;
        end
        
        %%% Position %%% e.g. NE
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'pos');
        if ~isempty(structMyArgument.value)
            myArg.stringPos  = structMyArgument.value;
            arguments        = structMyArgument.arguments;
        end
        %%%
        
        %%% Position in Percentage only applies if pos% is NOT a string!
        %%% e.g. pos% = [10,10] in %
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'pos%');
        if ~isempty(structMyArgument.value)
            myArg.stringPos            = structMyArgument.value;
            myArg.posPercentageByUser  = 1;
            arguments                  = structMyArgument.arguments;
        end
        %%%
        
        
        %%%
        
        %%% elementAnker only applies if pos is NOT a string! e.g. pos = [200,200]
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'posAnker');
        if ~isempty(structMyArgument.value)
            myArg.defaultPosAnker = structMyArgument.value;
            arguments             = structMyArgument.arguments;
        end
        
         %%% defaultOuterGapTileN %%% e.g. 10 //in px
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'gapN');
        if ~isempty(structMyArgument.value)
            myArg.defaultOuterGapTileN = structMyArgument.value;
            arguments            = structMyArgument.arguments;
        end
        %%%

        %%% defaultOuterGapTileS %%% e.g. 10 //in px
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'gapS');
        if ~isempty(structMyArgument.value)
            myArg.defaultOuterGapTileS = structMyArgument.value;
            arguments            = structMyArgument.arguments;
        end
        %%%

        %%% defaultOuterGapTileWE %%% e.g. 10 //in px
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'gapWE');
        if ~isempty(structMyArgument.value)
            myArg.defaultOuterGapTileWE = structMyArgument.value;
            arguments             = structMyArgument.arguments;
        end
        %%%

        %%% defaultWidthDivideTile %%% e.g. 40 %
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'divide');
        if ~isempty(structMyArgument.value)
            myArg.defaultWidthDivideTile = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%


        %%% Slidenumber %%% e.g. 'current' %
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'SlideNumber');
        if ~isempty(structMyArgument.value)
            myArg.defaultSlideNumber     = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%
        
                %%% SlideAddMethod %%% e.g. 'insert' or 'update' a slide %
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'SlideAddMethod');
        if ~isempty(structMyArgument.value)
            myArg.defaultSlideAddMethod = structMyArgument.value;
            arguments                   = structMyArgument.arguments;
        end

        %%% Height %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Height');
        if ~isempty(structMyArgument.value)
            myArg.userHeight                 = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%

        %%% Width %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Width');
        if ~isempty(structMyArgument.value)
            myArg.userWidth                  = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%

        %%% Width% %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Width%');
        if ~isempty(structMyArgument.value)
            myArg.defaultwidthPercentage = structMyArgument.value;
            arguments              = structMyArgument.arguments;
            myArg.widthPercentageByUser  = 1;
        end
        %%%

        %%% Width% %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Height%');
        if ~isempty(structMyArgument.value)
            myArg.defaultheightPercentage = structMyArgument.value;
            arguments              = structMyArgument.arguments;
            myArg.heightPercentageByUser  = 1;
        end
        %%%

        %%% Top %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Top');
        if ~isempty(structMyArgument.value)
            myArg.userTop      = structMyArgument.value;
            arguments          = structMyArgument.arguments;
        end
        %%%

        %%% Left %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Left');
        if ~isempty(structMyArgument.value)
            myArg.userLeft                   = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%
        
        
        %%% Absolute center position %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Left');
        if ~isempty(structMyArgument.value)
            myArg.userLeft                   = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%
        

    end
    
    if strcmp(out,'text')
        
        %% Gerneral Behaviour => 
        
        %% Setting up the page -  PageSetup
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'setPageFormat');
        
        if ~isempty(structMyArgument.value)
            myArg.pageFormat        = structMyArgument.value;
            arguments               = structMyArgument.arguments;
            myArg.isGeneralCommand  = 1;
            myArg.doSetPageFormat   = 1;
        end
        
        
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'setPageOrientation');
        
        if ~isempty(structMyArgument.value)
            myArg.pageOrientation        = structMyArgument.value;
            arguments                    = structMyArgument.arguments;
            myArg.isGeneralCommand       = 1;
            myArg.doSetPageOrientation   = 1;
        end
        
        % Adds a section
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'addSection');
        
        if ~isempty(structMyArgument.value)
            myArg.userSection            = structMyArgument.value;
            arguments                    = structMyArgument.arguments;
            myArg.isGeneralCommand       = 1;
            myArg.doAddSection           = 1;
        end
        
        
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'setFirstSlideNumber');
        if ~isempty(structMyArgument.value)
            myArg.firstSlideNumber  = structMyArgument.value;
            arguments               = structMyArgument.arguments;
            myArg.isGeneralCommand  = 1;
            myArg.doSetFirstSlideNumber   = 1;
        end
        
        % Do we want to apply a template - we should do this the first time
        % we call a new presentation
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'applyTemplate');
        
        if ~isempty(structMyArgument.value)
            myArg.templatePath      = structMyArgument.value;
            arguments               = structMyArgument.arguments;
            myArg.isGeneralCommand  = 1;
            myArg.doApplyTemplate   = 1;
        end
        
        % Do we want to open an existing presentation instead of creating
        % a new one or using an active one?
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'openExisting');
        
        if ~isempty(structMyArgument.value)
            myArg.existingPPTPath   = structMyArgument.value;
            arguments               = structMyArgument.arguments;
            myArg.isGeneralCommand  = 1;
            myArg.doOpenPPT         = 1;
        end
        
        
        
        % Do we want to save the presentation to the disk? => We will do
        % that in try catch way and inform the user when we where not able
        % to save the data
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'savePath');
        if ~isempty(structMyArgument.value)
            myArg.savePath          = structMyArgument.value;
            arguments               = structMyArgument.arguments;
            myArg.isGeneralCommand  = 1;
            myArg.doSavePPTPath     = 1;
        end
        
        % This only applies if savePath or saveFilename arument was set in
        % addition
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'savePassword');
        if ~isempty(structMyArgument.value)
            myArg.savePassword      = structMyArgument.value;
            arguments               = structMyArgument.arguments;
            myArg.hasSavePassword   = 1;
        end
        
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'close');
        if ~isempty(structMyArgument.value)
            myArg.close          = structMyArgument.value;
            arguments            = structMyArgument.arguments;
            myArg.isGeneralCommand  = 1;
            myArg.doClose           = 1;
        end
        
        
        % If the user wants to save the PPT not only to a desired path and
        % also wants to set a filename - otherwise the programm will set a
        % filename on its own - if we only set filename without savePPT =>
        % we will save the presentation to the active folder of matlab
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'saveFilename');
        if ~isempty(structMyArgument.value)
            myArg.saveFilename      = structMyArgument.value;
            arguments               = structMyArgument.arguments;
            myArg.isGeneralCommand  = 1;
            myArg.doSavePPTFilename = 1;
        end
        
        %%% Position %%% e.g. NE
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'pos');
        if ~isempty(structMyArgument.value)
            myArg.stringPos = structMyArgument.value;
            arguments        = structMyArgument.arguments;
        end
        %%%
        
        %%% Position in Percentage only applies if pos% is NOT a string!
        %%% e.g. pos% = [10,10] in %
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'pos%');
        if ~isempty(structMyArgument.value)
            myArg.stringPos            = structMyArgument.value;
            myArg.posPercentageByUser  = 1;
            arguments                  = structMyArgument.arguments;
        end
        %%%
        
        %%% elementAnker only applies if pos is NOT a string! e.g. pos = [200,200]
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'posAnker');
        if ~isempty(structMyArgument.value)
            myArg.defaultPosAnker = structMyArgument.value;
            arguments             = structMyArgument.arguments;
        end
        %%%

        %%% defaultOuterGapTileN %%% e.g. 10 //in px
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'gapN');
        if ~isempty(structMyArgument.value)
            myArg.defaultOuterGapTileN = structMyArgument.value;
            arguments            = structMyArgument.arguments;
        end
        %%%

        %%% defaultOuterGapTileS %%% e.g. 10 //in px
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'gapS');
        if ~isempty(structMyArgument.value)
            myArg.defaultOuterGapTileS = structMyArgument.value;
            arguments            = structMyArgument.arguments;
        end
        %%%

        %%% defaultOuterGapTileWE %%% e.g. 10 //in px
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'gapWE');
        if ~isempty(structMyArgument.value)
            myArg.defaultOuterGapTileWE = structMyArgument.value;
            arguments             = structMyArgument.arguments;
        end
        %%%

        %%% defaultWidthDivideTile %%% e.g. 40 %
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'divide');
        if ~isempty(structMyArgument.value)
            myArg.defaultWidthDivideTile = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%


        %%% Slidenumber %%% e.g. 'current' %
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'SlideNumber');
        if ~isempty(structMyArgument.value)
            myArg.defaultSlideNumber     = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%
        
        %%% SlideAddMethod %%% e.g. 'insert' or 'update' a slide %
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'SlideAddMethod');
        if ~isempty(structMyArgument.value)
            myArg.defaultSlideAddMethod = structMyArgument.value;
            arguments                   = structMyArgument.arguments;
        end
        %%%

        %%% Height %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Height');
        if ~isempty(structMyArgument.value)
            myArg.userHeight                 = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%

        %%% Width %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Width');
        if ~isempty(structMyArgument.value)
            myArg.userWidth                  = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%

        %%% Width% %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Width%');
        if ~isempty(structMyArgument.value)
            myArg.defaultwidthPercentage = structMyArgument.value;
            arguments              = structMyArgument.arguments;
            myArg.widthPercentageByUser  = 1;
        end
        %%%

        %%% Width% %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Height%');
        if ~isempty(structMyArgument.value)
            myArg.defaultheightPercentage = structMyArgument.value;
            arguments              = structMyArgument.arguments;
            myArg.heightPercentageByUser  = 1;
        end
        %%%

        %%% Top %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Top');
        if ~isempty(structMyArgument.value)
            myArg.userTop                    = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%

        %%% Left %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'Left');
        if ~isempty(structMyArgument.value)
            myArg.userLeft                   = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%

        %% Now get values for text

        %%% setTitle %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'setTitle');
        if ~isempty(structMyArgument.value)
            myArg.defaultTitle           = structMyArgument.value;
            arguments              = structMyArgument.arguments;
            myArg.doSetTitle             = 1;
        end
        %%%
        
        
        %%% asComment %%%
        structMyArgument        = deleteFromArgumentAndGetValue(arguments,'asComment');
        if ~isempty(structMyArgument.value)
            myArg.comment       = structMyArgument.value;
            arguments           = structMyArgument.arguments;
            myArg.addComment     = 1;
            %myArg.isGeneralCommand  = 1;
        end
        %%%

        %%% setBulletNumbers %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'setBulletNumbers');
        if ~isempty(structMyArgument.value)
            myArg.addBulletNumbers       = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        %%%

        %%% setBullets %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'setBullets');
        if ~isempty(structMyArgument.value)
            myArg.addBulletPoints        = structMyArgument.value;
            arguments              = structMyArgument.arguments;
        end
        
        %%% setTable %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'setTable');
        if ~isempty(structMyArgument.value)
            myArg.defaultTable         = structMyArgument.value;
            arguments                  = structMyArgument.arguments;
            myArg.doSetTable           = 1;
        end
        
        %%% Apply Tex interpreter %%%
        structMyArgument               = deleteFromArgumentAndGetValue(arguments,'TeX');
        if ~isempty(structMyArgument.value)
            myArg.doTexText            = structMyArgument.value;
            arguments                  = structMyArgument.arguments;
        end
        
        
        %%% setTableS %%% Same table on current slide if possible - if not
        %%% a new table will be added but just with dummy titles :-(
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'setTableS');
        if ~isempty(structMyArgument.value)
            myArg.defaultTable         = structMyArgument.value;
            arguments                  = structMyArgument.arguments;
            myArg.doSetTable           = 1;
            myArg.doSetTableS          = 1;
        end
        
        %%% setTable %%%
        structMyArgument       = deleteFromArgumentAndGetValue(arguments,'setHyperLink');
        if ~isempty(structMyArgument.value)
            myArg.defaultHyperlink     = structMyArgument.value;
            arguments                  = structMyArgument.arguments;
            myArg.doSetHyper           = 1;
        end
        
        
        %%% Do we want a visible frame arround the textBox?
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'hasFrame');
        
        if ~isempty(structMyArgument.value)
            myArg.addTextBoxFrame   = structMyArgument.value;
            arguments               = structMyArgument.arguments;
        end
        
        
        %%% Set color for frame arround the textBox
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'frameColor');
        
        if ~isempty(structMyArgument.value)
            myArg.defaultFrameColor   = structMyArgument.value;
            arguments                 = structMyArgument.arguments;
            myArg.addTextBoxFrame     = 1;
        end
        
        %%% Set width for frame arround the textBox
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'frameWidth');
        
        if ~isempty(structMyArgument.value)
            myArg.defaultFrameWidth   = structMyArgument.value;
            arguments                 = structMyArgument.arguments;
            myArg.addTextBoxFrame     = 1;
        end
        
        
        %%% Set type for frame arround the textBox
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'frameType');
        
        if ~isempty(structMyArgument.value)
            tempFrameType             = structMyArgument.value;
            
            % Get original title from type
            [isMember,indexFrameType] = ismember(tempFrameType,myArg.knownTextBoxFramesTitle);
            
            if ~isMember
                warning('The frame type is not known. Please check toPPT_config under toPPTText.knownTextBoxFramesTitle for all assignable frameTypes.')
            else
            
                myArg.defaultFrameType    = myArg.knownTextBoxFramesTitleOrg{indexFrameType};
                arguments                 = structMyArgument.arguments;
                
            end
            
            myArg.addTextBoxFrame     = 1;
        end
        
        %%% Set width for frame arround the textBox
        structMyArgument = deleteFromArgumentAndGetValue(arguments,'rotationAngle');
        
        if ~isempty(structMyArgument.value)
            myArg.defaultTextBoxRotation   = structMyArgument.value;
            arguments                      = structMyArgument.arguments;
        end
        

    end
    

    
    switch out
        case 'vec'
            
            %AutoGroupMin
            structMyArgument = deleteFromArgumentAndGetValue(arguments,'AutoGroupMin');
            if ~isempty(structMyArgument.value)
                
                if structMyArgument.value > myArg.defaultAutoGroupMin
                    myArg.defaultAutoGroupMin     = structMyArgument.value;
                end
                
                arguments                     = structMyArgument.arguments;
            end
            %
        case 'png'
            
            %%% magnify %%%
            structMyArgument       = deleteFromArgumentAndGetValue(arguments,'m');
            if ~isempty(structMyArgument.value)
                myArg.defaultMagnify         = structMyArgument.value;
                arguments              = structMyArgument.arguments;
            end
            %%%

    end
    
    myArg.externalParameters = '';
    
    if ~isempty(arguments)
        myArg.externalParameters = arguments; % Text to add in addText or just remaining arguments setted by user but not needed by this script at this point => we will just forward them
    end


end

% Local functions that are not used externally
function out = deleteFromArgumentAndGetValue(arguments,searchString)

value        = [];

try
    indexSearch = findStringInStructure(arguments,searchString);

catch myError
    indexSearch = -1;
    %Notfound => thats even better ;-)
end

if indexSearch ~= -1
    %%%Yes the user setted slidenumber to an other value => the value is
    %%%saved at slidenumber+1
    
    value = arguments{indexSearch+1}; %%Overwrite default value with user input
        
    %Now lets delete that value
    arguments = deleteFromStructure(arguments,[indexSearch,indexSearch+1]);
end

out.value       = value;
out.arguments   = arguments;
out.indexSearch = indexSearch;

end

function index = findStringInStructure(mystructure,mystring)

count = 0;
index = -1;

for ii=1:length(mystructure)
    try
    if strcmp(mystructure{ii},mystring)
        index(count+1) = ii;
        count = count+1;
    end
    catch
        % Seems not to be a string => No problem ;-)
    end
end

end

function newstructure = deleteFromStructure(structure,arrayIndices)

newstructure = '';

jj = 1;
for ii =1:length(structure)
    
    if isempty(find(arrayIndices==ii, 1))
        newstructure{jj} = structure{ii};
        jj = jj+1;
    end
    
end

end
