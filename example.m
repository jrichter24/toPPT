%% toPPT-2.1:
% * Author :  Jens Richter
% * eMail  :  jrichter@iph.rwth-aachen.de
% * Date   :  9th of June 2015


%% Adds all files of toPPT to matlab path
fullpath = mfilename('fullpath');
[pathstr1,~,~] = fileparts(fullpath);
addpath(genpath(pathstr1));

% Drive into toPPT folder.
cd(pathstr1);
clear pathstr1 fullpath;


%% Latest release notes (09 June 2015):
% * toPPT works together with *Matlab 2015a* and *Office 2013* (targeted system configuration).
% * Minor bug fixes.
% * Adding comments on slides.
% * Additional positioning system for figures and texts by pixel/percentage.
% * Minor changes in pptfigure to make it work together with *Matlab 2015a*
% * Adding of PowerPoint sections was introduced.
% * Possibility of using commandChains was added.
% * Load existing presentation and go on from arbitrary slide.
% * Change format and orientation of presentation.
% * Setting passwords when saving presentations.
% * Tables can use css tags and TeX.
% * Already existing fig-files can be used together with toPPT.


%% Older release nodes (28 February 2015):
% * toPPT works together with *Matlab 2014a* and *Office 2013*.
% * Minor bug fixes.
% * Added QR code ability.
% * Major changes in pptfigure to make it work together with *Office 2013*
% and *Matlab 2014b*
% * Moved all default values of toPPT into *toPPT_config.m*.
% * Updated used version of export_fig to to version published at 28 Feb 2015

%% Short description of toPPT:
% *toPPT* is a powerful tool for generating PowerPoint presentations
% programmatically defined in matlab.
% It will use different scripts to perform exports of figures,
% tables and texts. For this purpose it will use scripts written
% by myself and in addition scripts written by others *(see
% Acknowledgment)*.


%% Acknowledgment:
% 
% # *export_fig-script* by Oliver Woodford
    % http://www.mathworks.com/matlabcentral/fileexchange/23629-exportfig
    % => Is used to generate pictures (png) from the desired figures.
% # *pptfigure.m* by Dmitriy Aronov
    % http://www.mathworks.com/matlabcentral/fileexchange/30124-smart-powerpoint  -exporter/content/pptfigure.m
    % => Is used for generating editable figures. An extracted part of it
    % is used as tex2ppt converter.
% # *RGB triple of color name*, Version 2 by Kristjan Jonasson
    % http://www.mathworks.com/matlabcentral/fileexchange/24497-rgb-triple-of-color-name--version-2
    %  => Is used to transform color names to hex values.   
% # *Edit Distance Algorithm* by  Reza Ahmadzadeh
    % http://www.mathworks.com/matlabcentral/fileexchange/39049-edit-distance-algorithm
    % => Is used to calculate the edit distance for placing a new slide near
    % an other slide with a certain title.
% # *QR Code Generator based on zxing* taken from another project of myself
    % http://www.mathworks.com/matlabcentral/fileexchange/49808-qr-code-generator-based-on-zxing
    % => Is used to create the QR-Codes (Please also check the Acknowledgment of QR-Code directly).
    
%% To make toPPT work:
% Simply add toPPT and all subdirectories and folders to your matlab path.
% Example.m will do this automatically for you. 
% The path has to be writable and readable because toPPT has to
% temporally save data.
% For a detailed help simply use *'help toPPT'*.


%% Examples - Introduction:
% With some examples we will learn how toPPT works - In addition
% we will use a little subscript _example_helper.m_ to create dummy figures.


%% BASIC COMMANDS
%
%% Preparation:
% This example presentation is optimized for a page format of 16:9.
% Refer to *Example 12* for more information.
toPPT('setPageFormat','16:9');


%% Example 1 - Adding a figure:
% The challenge is to place an image in png format centred in your
% presentation.
% For this example PowerPoint is already closed or open. The program will automatically use an open presentation for adding
% content. If PowerPoint is closed it will open it and create a blank presentation.
 
figure1 = example_helper(1); %Gets a figure
toPPT(figure1); % By default a new slide is created when adding a figure


%% Example 2 - Setting a title:
% We want to add a title to the slide of Example 1.

toPPT('setTitle','My first figure with toPPT-beta1, Example 1-2');


%% Example 3a - Adding a figure at a certain position:
% We want to add an image in png format to the upper right half _(NorthEastHalf)_ 
% of the presentation and one image to the lower half _(SouthEastHalf)_ of the image.
% For detailed information use *'help getPosParameters'*.


figure2 = example_helper(2);

toPPT(figure1,'pos','NEH'); % NEH = NorthEastHalf
toPPT(figure2,'pos','SEH','SlideNumber','current'); % NEH = SouthEastHalf

toPPT('setTitle','Example 3a - Adding a figure at a certain position');

%% Example 3b - Changing the height and width of a figure:
% We want to add an image to the middle tile of the slide and stretch it to 100% of the
% width.
% We also want to add another image and stretch it to 200px in height
% which will be placed in the _Middle East_.


toPPT(figure1,'pos','M','Width%',100,'Height%',100);
toPPT('setTitle','Example 3b - Defining the height and width of a figure in percentage');

toPPT(figure1,'pos','ME','Height',200,'Width',100); % ME = MiddelEast
toPPT('setTitle','Example 3b - Defining the height and width of a figure in pixel');


%% Example 3c - Adding a figure at a certain position by specifying position in percentage:
% We want to add an image in png format and specify the position by coordinates 
% related to the dimension of the current presentation.
% In addition we specify a 'posAnker'. This means the point specified by
% 'pos%' is the anker point of the current image. In case only 'pos' is
% used (without %) together with an numerical array the values assumed to
% be pixel coordinates.
% For detailed information use *'help getPosParameters'* and *'help toPPT'*.


posPercentageX = 75; % 75% of the width of the current presentation
posPercentageY = 55; % 55% of the height of the current presentation
toPPT(figure1,'pos%',[posPercentageX,posPercentageY],'Height%',30,'posAnker','NW');
toPPT(figure2,'pos%',[posPercentageX,posPercentageY],'Height%',30,'posAnker','SW','SlideNumber','current');
toPPT(figure2,'pos%',[20,55],'Width%',30,'posAnker','C','SlideNumber','current');

% When specifying Width and Height the aspect ratio is not necessarily kept
toPPT(figure1,'pos%',[55,55],'Height%',40,'Width%',20,'posAnker','C','SlideNumber','current'); 

toPPT('setTitle','Example 3c - Adding a figure at a certain position by specifying position in pixel or percentage');

%
posPixelX = 20;
posPixelY = 300;
toPPT(figure1,'pos',[posPixelX,posPixelY],'Height%',30,'posAnker','NW');
toPPT(figure2,'pos',[posPixelX+200,posPixelY],'Height',200,'posAnker','NW','SlideNumber','current');
toPPT(figure2,'pos',[posPixelX+500,posPixelY],'Height%',40,'posAnker','NW','SlideNumber','current');
toPPT(figure1,'pos',[posPixelX+500,posPixelY],'Width%',20,'posAnker','SW','SlideNumber','current');
toPPT(figure1,'pos',[posPixelX,posPixelY],'Height%',15,'Width%',20,'posAnker','SW','SlideNumber','current');

toPPT('setTitle','Example 3c - Adding a figure at a certain position by specifying position in pixel or percentage');


%% Example 3d - Adding an existing figure:
% The challenge is to load an existing figure and add it to our
% presentation. The path can be defined relative or absolute.

figPath = [pwd,'\testLoadMe.fig'];
toPPT('existingFigure',figPath); % By default a new slide is created when adding a figure

toPPT('setTitle','Example 3d - Adding an existing figure');


%% Example 3e  - Changing the padding/gap of the active placing area of the slide:
% We want to change the default gap between our centered tile in the North and in 
% the West/East. 


toPPT(figure1,'gapN',200,'gapWE',150);
toPPT('setTitle','Example 8a  - Changing the padding/gap of the active placing area of the slide');

%% Example 3f - Change the 'dividing' of the slide:
% We want to change the divide property of the tiles - by default each
% image is stretched to 100% of the tile height by keeping the aspect
% ratio. For more information see *'help toPPT'*.


toPPT(figure1,'pos','E','Width%',60,'divide',30);
toPPT(figure2,'pos','W','SlideNumber','current','Width%',60,'divide',30);
toPPT('setTitle','Example 8b - Change the dividing of the slide');



%% Example 4 - Adding a text with a bullet point:
% We want to set a title and add a text as a bullet point.
% *By default text is added centered to the current slide as a bullet point*
% (Refers to the command _'SlideNumber','current'_).
% *Remember: By default figures are added to a new slide*
% (Refers to the command _'SlideNumber','append'_).

toPPT('My first bullet item','pos','W','SlideNumber','append'); % We want to add it to the W = West
toPPT('setTitle','Example 4 - Adding a text with a bullet point');


%% Example 4a - Adding a text with a bullet point to a new slide:
% We want to add a list of texts with bullet points to a *NEW* slide. 
% *It is important to put the list into a cell aray {}.*

toPPT({'Text1','Text2','Text3'},'SlideNumber','append');
toPPT('setTitle','Example 4a - Adding a text with a bullet point to a new slide');


%% Example 4b - Adding multiple texts with a bullet points to a new slide:
% We add a list of texts in the _East_ (new slide) with numbers instead of
% bullet points.

toPPT({'Text1','Text2','Text3'},'SlideNumber','append','setBulletNumbers',1,'pos','W');
toPPT({'Text1','Text2','Text3'},'setBullets',0,'pos','E'); % E = East
toPPT('setTitle','Example 4b - Adding multiple texts with a bullet point to a new slide');

%% Example 4c - Applying bold, underlined and italic to text elements (css-tags):
% We want to make some parts of our texts underlined, italic or/and bold.

toPPT({'<b>Bold</b>','<b><u>Bold and underlined</u></b> meets <i>Italic</i>','Normal'},'SlideNumber','append');
toPPT('setTitle','Example 4c - Applying bold, underlined and italic to text elements (css-tags)');

%% Example 4d - Changing the color, size and font of text elements (css-tags):
% We want to assign another color, different size and a different font to
% our texts.
% Have a look on _*'toPPT accepts the following predefined colors'*_ in this example-file for a list of all known colors.

toPPT({'<s color:red>Red text,</s> Black text, <s color:blue>Blue text</s>','<s font-family:Times New Roman>Using Times New Roman</s>','<s font-size:40>BIG</s>'},'SlideNumber',...
    'append');
toPPT('setTitle','Example 4d - Changing the color, size and font of text elements (css-tags)');

%% Example 4e - Changing the color, size and font and apply bold, underlined and italic to text elements:
% Now we want to combine different tags (<s>,<u>...) together.

toPPT({'<s color:green; font-size:36; font-family:Aharoni>Mixed One</s>','<i><u><b><s font-family:Times New Roman>Bold, underlined,italic and Times</s></b></u></i>'},...
    'SlideNumber','append');
toPPT('setTitle','Example 4e - Changing color, size and font and apply bold, underlined and italic');

%% Example 4f -  Using Hex-Code for defining the color of a text element:
% Now we want to use hex colors - *it is important to use a leading '#' in front of
% the hexcolor.*
% The syntax for the tag *s* (s means style) is adapted from the css
% language. It is absolutely necessary that each argument within the *s*-tag is
% closed via ";" (for the last argument the |>| works as closing ";")

toPPT('<s bg:yellow; color:#5F9EA0; font-size:22; font-family:Aharoni>CadetBlue</s> meets <s color:#FFA07A; font-size:22; font-family:Aharoni>LightSalmon</s>',...
    'SlideNumber','append');
toPPT('setTitle','Example 4f -  Using Hex-Code for defining the color of a text element');


%% Example 4 $\eta$ -  Using TeX in texts:
% Now we want to inspect how the TeX interpreter works. In general the TeX
% interpreter is turned on by default. Example 17 _*Saving a presentation*_ shows that it is sometimes
% necessary to turn off the interpreter.
% An overview of all possible TeX characters can be found in _*"toPPT accepts the following predefined TeX parameters"*_ in this example-file.
% In general _TeX_ can be combined with all other css-tags.

toPPT({'Lets start with something easy x = \sqrt\x^2','I like red formulas <s color:red>x = \sqrt\x^2<\s>',...
    'And big blue ones <i><b><s color:blue; font-size:40; font-family:Times New Roman>x = \sqrt\x^2<\s></b></i>'},'SlideNumber','append');
toPPT('setTitle','Example 4 \eta -  Using TeX in texts');

%% Example 4 $\gamma$ -  Using TeX and css-tags in texts at the same time:
% It is also possible to use the TeX tag \color
% By default the css tags (s etc.) have priority in combination with _TeX_
% tags.

toPPT('Colored formula with Tex: \color{red}a^2+b^2 = c^2','SlideNumber','append');
toPPT('TeX color is overwritten by "normal" tag: <s color:blue>\color{red}a^2+b^2 = c^2</s>','pos','E');
toPPT('setTitle','Example 4 \gamma -  Using TeX and css-tags in texts at the same time');


%% Example 4 $\phi$ - Using TeX in texts together with matlab build in texlabel-function:
% We also want to do some more complicated stuff, sometimes we are to
% lazy to write the _TeX_-code so we can use the inbuild matlab function
% _texlabel_.

myTexText = texlabel('lambda12^(3/2)/pi - pi*delta^(2/3)');

toPPT(myTexText,'SlideNumber','append');
toPPT(['<b><s bg:orange; color:green>',texlabel('H(x)*psi(x)=E*psi(x)'),'<\s><\b>'],'pos','NEH');
toPPT({'TeX Interpreter is off:',['<s color:blue>',texlabel('H(x)*psi(x)=E*psi(x)'),'<\s>']},'pos','SWH','TeX',0); % TeX interpreter turned off

toPPT('setTitle','Example 4 \phi - Using TeX in texts together with matlab build in texlabel-function');


%% Example 5a - Updating an existing slide identified by its title:
% The parameter _SlideNumber_ can be used differently:
%
% * A slide can be found by its title.
% * Even if the spelling is slightly wrong the slide can be found via edit distances.

toPPT('<s color:blue>Check "Example 4f" for updated texts.<\s>','SlideNumber','append'),
toPPT('setTitle','Example 5a - Updating an existing slide identified by its title');

toPPT('<s color:blue>New updated Text on slide "Example 4f"<\s>','SlideNumber',...
    'Example 4f -  Using Hex-Code for defining the color of a text element','pos%',[1,50],'setBullets',0); % Found by its title.

toPPT('<s color:red>New updated Text on slide "Example 4f" with wrong spelling of title<\s>',...
    'SlideNumber','Aaxmple 3f -  Uesin HeX-Kode for efining the cooulor of a text ele','pos%',[1,60],'setBullets',0,'Width%',90); % Found by its title with wrong spelling.

%% Example 5b - Inserting content after an existing slide identified by its title:
% The parameter _SlideAddMethod_ can be used to insert content after an
% existing slide instead of updating the slide. *By default slides are
% updated.*

toPPT('<s color:blue>Check slide after "Example 4e" for inserted text.<\s>','SlideNumber','append'),
toPPT('setTitle','Example 5b - Inserting content after an existing slide identified by its title');

toPPT('New inserted Text after Slide of Example 4e','SlideNumber','Example 4e - Changing color, size and font and apply bold, underlined and italic'...
    ,'SlideAddMethod','insert','pos','NW'); % Inserting content after choosen slide

toPPT('setTitle','Example 4eb - Inserting content after an existing slide identified by its title');

%% Example 5c - Updating an existing slide identified by its slide number:

toPPT('<s color:blue>Check first slide for updated text.<\s>','SlideNumber','append'),
toPPT('setTitle','Example 5c - Updating an existing slide identified by its slide number');

toPPT('<b><i><s font-family:Times New Roman;color:blue;font-size:30>Example toPPT() 2.1 by Jens Richter<\s><\i><\b>','SlideNumber',1,...
    'setBullets',0,'Width%',25,'pos%',[1,50],'posAnker','NW'); % Adds a new textbox to slide 1.

%% Example 5d - Inserting content in a new slide at slide number 4 (will shift all following slides):

toPPT('<s color:blue>Check slide after slide 8 for inserted slide 9.<\s>','SlideNumber','append'),
toPPT('setTitle','Example 5d - Inserting content in a new slide after slide number 8 (will shift all following slides)');

toPPT('New inserted Text <b>AFTER</b> slide Number 8','SlideNumber',8,'SlideAddMethod','insert','pos','NW'); % New slide will become slide 9
toPPT('setTitle','Example 5d - Inserting content in a new slide at slide number 4 (will shift all following slides)');


%% Example 6 - Adding Sections
% We want to add a section to our presentation. This can be used to subdivide
% your presentation into different thematic sections.
% For a list of all available parameters call *help toPPT*. 

toPPT({'<b>Add a section for the basic commands!<\b>',...
    '<b>Add a section for texts and tags.<\b>',...
    '<b>Add a section for updating and inserting slides.<\b>',...
    '<b>Add a section for Special commands.<\b>'},'SlideNumber','append');
toPPT('setTitle','Example 6 - Adding Sections');

toPPT('addSection','Basic commands','SlideNumber','My first figure with toPPT-beta1, Example 1-2','SlideAddMethod','before');
toPPT('addSection','Texts, TeX and tags','SlideNumber','Example 4 - Adding a text with a bullet point','SlideAddMethod','before');
toPPT('addSection','Inserting, updating slides','SlideNumber','Example 5a - Updating an existing slide identified by its title','SlideAddMethod','before');
toPPT('addSection','Sections and Templates','SlideNumber','Example 6 - Adding sections to the presentation','SlideAddMethod','before');



%% Example 7 - Applying a template to the active presentation:
% Now we want to apply a template to our presentation. *The template file
% path has to be absolute*.

templatePath = [pwd,'\IPH Slides_EN.potx']; % The template path has to be absolute!
toPPT('applyTemplate',templatePath);

toPPT('<b>We applied a template!<\b>','SlideNumber','append');
toPPT('setTitle','Example 7 - Applying a template to the active presentation');



%% Example 8 - Adding Tables:
% We want to create a new slide and add a centered table to this slide. In
% addition we want to add two smaller tables on another slide positioned in the West and East.
%
% * All content for the table is put into cell arrays *(One cell for the captions and one cell/matrix etc. for the content)*. 
% It doesn't matter if the content-cell has numeric and string values at the same time.
% toPPT will automatically detects if it has to _rotate_ the content to fit the number of caption entries.
%
% * In general setTable accepts different arguments as Input:
%
% # *A two dimensional cell* that fully defines the desired cell *NOT* including the
% captions. (8a)
% # *Row or column vectors.* 
% toPPT will automatically detect if it is a row or column vector by checking 
% the number of caption entries. 
% It is *NOT* possible to add row and column vectors at the same time. (8b)
% # *A numeric matrix.* (8c)
%
% The syntax is: *toPPT('setTable',{stringCellTableCaption, matrix/Vector/cell})*
% UPDATE: Each cell can accept s-tags and TeX!

helpTableCell      = cell(2,3);
helpTableCell{1,1} = '<s bg:red; color:white><i>Me is first<\i><\s>';
helpTableCell{1,2} = 'Me is second';
helpTableCell{1,3} = ['<b><s color:orange>','Me is formula: ',texlabel('H(x)*psi(x)=E*psi(x)'),'<\s><\b>'];
helpTableCell{2,1} = '<s color:#5F9EA0; font-size:22; font-family:Aharoni>Me is fourth</s>';
helpTableCell{2,2} = 'Me is fifth';
helpTableCell{2,3} = 'Me is sixth';

% First slide -  It is important to add "helpTableCell" into an extra cell {}
toPPT('setTable',{{'<b>Caption1<\b>','<b>Caption2<\b>','<b>Caption3<\b>'},{helpTableCell}},'SlideNumber','append','Width',400,'Height',150,'pos%',[48,55],'posAnker','E');

% Second slide - auto rotate will be used
toPPT('setTable',{{'<b>Caption1<\b>','<b>Caption2<\b>'},{helpTableCell}},'SlideNumber','current','Width',400,'Height',150,'pos%',[52,55],'posAnker','W');
toPPT('setTitle','Example 8a - Adding tables using cell of strings (right table is using auto rotate)');

% Third slide - It is important to add all vectors into an extra cell {}
helperVector1 = [1,2,3];
helperVector2 = {'Me is 1','Me is 2','Me is 3'};

toPPT('setTable',{{'Col 1','Col 2'},{helperVector1,helperVector2}},'SlideNumber','append','pos','ME');

% With autorotation
toPPT('setTable',{{'Col 1','Col 2','Col 3'},{helperVector1,helperVector2}},'SlideNumber','current','pos','MW');
toPPT('setTitle','Example 5b - Adding tables using vectors and cells mixed (right table is using auto rotate)');

% Fourth slide
toPPT('setTable',{{'Col 1','Col 2'},ones(2,2)},'SlideNumber','append','pos','E','Width',200,'Height',150);
toPPT('setTable',{{'Col 1','Col 2'},ones(3,2)},'pos','W','Width',200,'Height',150);
toPPT('setTitle','Example 5c - Adding tables using a matrix');

%% Adding a section for tables
toPPT('addSection','Tables','SlideNumber','Example 8a - Adding tables using cell of strings (right table is using auto rotate)','before');

%% Example 9 - Deactivating toPPT temporally:
% Maybe we don't want to add something to the presentation all the time
% because we are just testing our script...

doPublishing = 0;
toPPT(figure1,'pub',doPublishing); % Nothing happens

doPublishing = 1;
toPPT(figure1,'pub',doPublishing); % Now we publish
toPPT('setTitle','Example 9 - Deactivating toPPT temporally and reactivating');

%% Adding a section for Special commands II
toPPT('addSection','Special commands II','SlideNumber','Example 9 - Deactivating toPPT temporally and reactivating','before');

%% Example 10 - Adding comments:
% We want to add a comment.
% For a list of all available parameters call *help toPPT*. 

toPPT({'Some text 1','Some text2'},'SlideNumber','append');

author = 'Jens Richter';
initials = 'JR';
comment = 'I have to make some annotations!';

toPPT('asComment',{author,initials,comment});
toPPT('setTitle','Example 10 - Adding comments');

%% Example 11 - Using commandChains:
% CommandChains can be used to design a presentation e.g. within in other script file. 
% The advantage is that no escaping is necessary and the presentation can be easily bundled. 
% Annotation: In practice big CommandChains can exhibit the available physical memory. 
% Especially when a lot of figures are put into the commandChain.
% For a list of all available parameters call *help toPPT*. 

commandsCell{1} = {figure1,'Width%',50,'pos','ME'};
commandsCell{2} = {'setTable',{{'<b>Caption1<\b>','<b>Caption2<\b>','<b>Caption3<\b>'},{helpTableCell}},'SlideNumber','current','Width',400,'Height',150,'pos%',[48,55],'posAnker','E'};
commandsCell{3} = {'setTitle','Example 11 - Using commandChains'};

for ii=1:numel(commandsCell)
    toPPT('commandChain',commandsCell{ii});
end



%% Example 12 - Change format and orientation of presentation:
% In case the presentation needs to have another orientation or format this
% option can be used. 
% Annotation: It is recommended to this before adding any content to the slide.
% For a list of all available parameters call *help toPPT*. 

toPPT({'Changing the page orientation is done e.g by <b>toPPT(''setPageOrientation'',''landscape'');</b>','',...
   'Changing the page orientation is done e.g by <b>toPPT(''setPageFormat'',''16:9'');</b>'},'SlideNumber','append','setBullets',0,'pos%',[1,20],'Width%',80);
toPPT('setTitle','Example 12 - Change format and orientation of presentation');



%% Example 13 - Change the resolution of the figure for exporting:
% We want to *increase* (for publications etc.)/ *decrease* (smaller file size, faster)
% the resolution of the png image that is used for exporting a figure.

toPPT(figure1,'m',3); % Higher value, default is m = 2
toPPT('setTitle','Example 13 - Change the resolution of the figure for exporting (high resolution)');

%% Adding a section for 'Figure editing and figure generation'
toPPT('addSection','Figure editing and figure generation','Example 13 - Change the resolution of the figure for exporting (high resolution)','before');

toPPT(figure1,'m',1); % Lower value, default is m = 2
toPPT('setTitle','Example 13 - Change the resolution of the figure for exporting (low resolution)');



%% Example 14 - Export figure as vector graphic:
% We want to export our image as vector graphic. This process is quite
% challenging. For huge figures with a lot of points it is not recommended
% to use this function. *In case the export fails toPPT will use the default
% png export as a fall back solution anyway.*

figureVector = example_helper(2);
toPPT(figureVector,'format','vec');
toPPT('setTitle','Example 14 - Export figure as vector graphic');


%% Example 15 - Specify parameters for export_fig directly via toPPT:
% toPPT makes strong use of *'export_fig'*. Maybe in some situations it is
% necessary to supply parameters for export_fig directly. For this reason
% all commands that are not recognized be toPPT are directly forwarded as
% input to export_fig. Please refer to *help export_fig*. For some commands
% export_fig needs to have Ghostscript installed.
%
% In this example we change the colorspace to CMYK and gray for the same
% figure before exporting to powerpoint.


toPPT(figure1,'pos','W','-CMYK'); % colorspace CMYK
toPPT(figure1,'pos','E','SlideNumber','current','-gray'); % colorspace gray
toPPT('setTitle','Example 15 - Change the colorspace (specify parameters for exportfig directly via toPPT)');


%% Example 16a - QR-Code with test message 'Hello world':
% *Annotation:* QR-Code generation is only available for *Matlab2014a* and
% later. The QR-Code generation is based on *zxings* java libraries. 
% For this purpose toPPT will import all
% necessary  jar files on the fly from a maven repository server (check
% *qrcode_config.m* for details). In addition QRcode_gen has an option to 
% download the necessary jar files and save them into a predefined directory. 
% If you are not willing to download precompiled jars from the web 
% (what is understandable in terms of security issues) you can access
% zxings open source project (http://github.com/zxing) check the source code
% and compile the jars yourself. For a list of all available parameters call *help toPPT*. 
%
% We want to simply add a QR-Code with a test message.

message = 'This is a hello world test message';
toPPT(message,'TextAsQR',1,'SlideNumber','append');
toPPT('setTitle','Example 16a - QR-Code with test message "Hello world"');

%% Example 16b - Adding MECard as QR Code
% We want to add an QR code to the left corner at the bottom. The QR code
% in this example is representing a MECARD. In addition we want to change
% the background and QR-Code color.

message = 'MECARD:N:Skywalker,Luke;ADR:76 9th Avenue, 4th Floor, New York, NY 10011;TEL:1234567891011;EMAIL:luke@skywalker.com;;';

toPPT(message,'TextAsQR',1,'QR-Version',40,'pos','SW','Left',5,'gapS',15,'QR-Color','#00549F','QR-BgColor','#D5FFFF','SlideNumber','append');
toPPT(['<s color:#00549F; font-size:22; font-family:Aharoni>Thank you for your attention.',char(13),...
    'Please scan the QR Code for Luke Skywalkers contact information.</s> '],'pos','ME','setBullets',0);
toPPT('setTitle','Example 16b - Adding MECard as QR Code');


%% Example 16c - Downloading the necessary jars from the internet:
% This will just add the necessary syntax for downloading the jars. We will
% not do that at this point so you can decide on your own if you want to
% download them.

toPPT('Downloading jars for QR-Code generator is done by <b>toPPT(someStringMessage,''QR-DownloadJars'',1)</b>','SlideNumber','append');
toPPT('setTitle','Example 16c - Syntax for downloading the necessary jars from the internet');




%% Example 17 - Saving a presentation:
% Now we want to save our presentation.
% There are three ways of saving
%
% _Case 1:_ We only supply the _*savePath*_ as argument
%
%     => Presentation will be saved in path as 'The Current Date and Time
%     _new_presentation.pptx'.
%
% _Case 2:_ We only supply the _*saveFilename*_ as argument
%
%     => Presentation will be saved in pwd with the desired filename.
%
% _Case 3:_ We supply _*savePath*_ and _*saveFilename*_ as argument
%
%     => Presentation will be saved in savePath with the desired filename.
%
% *The savepath has to be absolute.*

% * Setting passwords when saving presentations.

savePath = pwd;
saveFilename = 'My first presentation with toPPT';

toPPT('savePath',savePath,'saveFilename',saveFilename);
toPPT(['We saved our presentation to: ',savePath,'/',saveFilename]...
    ,'SlideNumber','append','TeX',0); % Because we want to "show" a filename we should turn of TeX otherwise all "\" will be gone etc.

toPPT('setTitle','Example 17 - Saving a presentation');

%% Adding a section for saving, loading and closing presentation
toPPT('addSection','Saving, loading and closing presentation','SlideNumber','Example 17 - Saving a presentation','before');

%% Example 19 - Saving a presentation password protected
% Saving a presentation password protected can be simply done by adding the
% attriute 'savePassword' and the password itself as argument.
% Annotation: A password protected presentation can NOT be loaded via
% toPPT.

toPPT('Saving a presentation password protected is done by <b>toPPT(''savePath'',''YOURSAVEPATH'',''saveFilename'',''YOURSAVEFILENAME'',''savePassword'',''YOURSAVEPASSWORD'',);</b>',...
    '','SlideNumber','append','setBullets',0,'pos%',[1,20],'Width%',95);
toPPT('setTitle','Example 19 - Saving a presentation password protected');


%% Example 20 - Loading an existing presentation
% In case an already existing presentation should be continued we can load
% this presentation via toPPT in case no password was set previously.

toPPT({'Open an existing presentation (absolute Path) is done by <b>toPPT(''openExisting'',''C:\myPresentation.pptx'');</b>','',...
   'Open an existing presentation (relative Path) is done by <b>toPPT(''openExisting'',''.\myPresentation.pptx'');</b>'},...
   'SlideNumber','append','setBullets',0,'pos%',[1,20],'Width%',95,'TeX',0);

toPPT('setTitle','Example 10 - Loading an existing presentation');


%% Example 21 - Closing a presentation:
% We can close a presentation with toPPT - this can be helpful if we want
% to create multiple different presentations and close them after we saved 
% them. Because we do not want to close the presentation in this example we
% just explain it on an extra slide.

toPPT('Closing a presentation is done via the syntax <b>toPPT(''close'',1)</b>','SlideNumber','append');
toPPT('setTitle','Example 21 - Closing a presentation');

%% Done with the example. Let’s close all generated figures.
close all;


%% toPPT accepts the following predefined colors:
%     %White colors
%     'FF','FF','FF', 'White'
%     'FF','FA','FA', 'Snow'
%     'F0','FF','F0', 'Honeydew'
%     'F5','FF','FA', 'MintCream'
%     'F0','FF','FF', 'Azure'
%     'F0','F8','FF', 'AliceBlue'
%     'F8','F8','FF', 'GhostWhite'
%     'F5','F5','F5', 'WhiteSmoke'
%     'FF','F5','EE', 'Seashell'
%     'F5','F5','DC', 'Beige'
%     'FD','F5','E6', 'OldLace'
%     'FF','FA','F0', 'FloralWhite'
%     'FF','FF','F0', 'Ivory'
%     'FA','EB','D7', 'AntiqueWhite'
%     'FA','F0','E6', 'Linen'
%     'FF','F0','F5', 'LavenderBlush'
%     'FF','E4','E1', 'MistyRose'
%     %Grey colors'
%     '80','80','80', 'Gray'
%     'DC','DC','DC', 'Gainsboro'
%     'D3','D3','D3', 'LightGray'
%     'C0','C0','C0', 'Silver'
%     'A9','A9','A9', 'DarkGray'
%     '69','69','69', 'DimGray'
%     '77','88','99', 'LightSlateGray'
%     '70','80','90', 'SlateGray'
%     '2F','4F','4F', 'DarkSlateGray'
%     '00','00','00', 'Black'
%     %Red colors
%     'FF','00','00', 'Red'
%     'FF','A0','7A', 'LightSalmon'
%     'FA','80','72', 'Salmon'
%     'E9','96','7A', 'DarkSalmon'
%     'F0','80','80', 'LightCoral'
%     'CD','5C','5C', 'IndianRed'
%     'DC','14','3C', 'Crimson'
%     'B2','22','22', 'FireBrick'
%     '8B','00','00', 'DarkRed'
%     %Pink colors
%     'FF','C0','CB', 'Pink'
%     'FF','B6','C1', 'LightPink'
%     'FF','69','B4', 'HotPink'
%     'FF','14','93', 'DeepPink'
%     'DB','70','93', 'PaleVioletRed'
%     'C7','15','85', 'MediumVioletRed'
%     %Orange colors
%     'FF','A5','00', 'Orange'
%     'FF','8C','00', 'DarkOrange'
%     'FF','7F','50', 'Coral'
%     'FF','63','47', 'Tomato'
%     'FF','45','00', 'OrangeRed'
%     %Yellow colors
%     'FF','FF','00', 'Yellow'
%     'FF','FF','E0', 'LightYellow'
%     'FF','FA','CD', 'LemonChiffon'
%     'FA','FA','D2', 'LightGoldenrodYellow'
%     'FF','EF','D5', 'PapayaWhip'
%     'FF','E4','B5', 'Moccasin'
%     'FF','DA','B9', 'PeachPuff'
%     'EE','E8','AA', 'PaleGoldenrod'
%     'F0','E6','8C', 'Khaki'
%     'BD','B7','6B', 'DarkKhaki'
%     'FF','D7','00', 'Gold'
%     %Brown colors
%     'A5','2A','2A', 'Brown'
%     'FF','F8','DC', 'Cornsilk'
%     'FF','EB','CD', 'BlanchedAlmond'
%     'FF','E4','C4', 'Bisque'
%     'FF','DE','AD', 'NavajoWhite'
%     'F5','DE','B3', 'Wheat'
%     'DE','B8','87', 'BurlyWood'
%     'D2','B4','8C', 'Tan'
%     'BC','8F','8F', 'RosyBrown'
%     'F4','A4','60', 'SandyBrown'
%     'DA','A5','20', 'Goldenrod'
%     'B8','86','0B', 'DarkGoldenrod'
%     'CD','85','3F', 'Peru'
%     'D2','69','1E', 'Chocolate'
%     '8B','45','13', 'SaddleBrown'
%     'A0','52','2D', 'Sienna'
%     '80','00','00', 'Maroon'
%     %Green colors
%     '00','80','00', 'Green'
%     '98','FB','98', 'PaleGreen'
%     '90','EE','90', 'LightGreen'
%     '9A','CD','32', 'YellowGreen'
%     'AD','FF','2F', 'GreenYellow'
%     '7F','FF','00', 'Chartreuse'
%     '7C','FC','00', 'LawnGreen'
%     '00','FF','00', 'Lime'
%     '32','CD','32', 'LimeGreen'
%     '00','FA','9A', 'MediumSpringGreen'
%     '00','FF','7F', 'SpringGreen'
%     '66','CD','AA', 'MediumAquamarine'
%     '7F','FF','D4', 'Aquamarine'
%     '20','B2','AA', 'LightSeaGreen'
%     '3C','B3','71', 'MediumSeaGreen'
%     '2E','8B','57', 'SeaGreen'
%     '8F','BC','8F', 'DarkSeaGreen'
%     '22','8B','22', 'ForestGreen'
%     '00','64','00', 'DarkGreen'
%     '6B','8E','23', 'OliveDrab'
%     '80','80','00', 'Olive'
%     '55','6B','2F', 'DarkOliveGreen'
%     '00','80','80', 'Teal'
%     %Blue colors
%     '00','00','FF', 'Blue'
%     'AD','D8','E6', 'LightBlue'
%     'B0','E0','E6', 'PowderBlue'
%     'AF','EE','EE', 'PaleTurquoise'
%     '40','E0','D0', 'Turquoise'
%     '48','D1','CC', 'MediumTurquoise'
%     '00','CE','D1', 'DarkTurquoise'
%     'E0','FF','FF', 'LightCyan'
%     '00','FF','FF', 'Cyan'
%     '00','FF','FF', 'Aqua'
%     '00','8B','8B', 'DarkCyan'
%     '5F','9E','A0', 'CadetBlue'
%     'B0','C4','DE', 'LightSteelBlue'
%     '46','82','B4', 'SteelBlue'
%     '87','CE','FA', 'LightSkyBlue'
%     '87','CE','EB', 'SkyBlue'
%     '00','BF','FF', 'DeepSkyBlue'
%     '1E','90','FF', 'DodgerBlue'
%     '64','95','ED', 'CornflowerBlue'
%     '41','69','E1', 'RoyalBlue'
%     '00','00','CD', 'MediumBlue'
%     '00','00','8B', 'DarkBlue'
%     '00','00','80', 'Navy'
%     '19','19','70', 'MidnightBlue'
%     %Purple colors
%     '80','00','80', 'Purple'
%     'E6','E6','FA', 'Lavender'
%     'D8','BF','D8', 'Thistle'
%     'DD','A0','DD', 'Plum'
%     'EE','82','EE', 'Violet'
%     'DA','70','D6', 'Orchid'
%     'FF','00','FF', 'Fuchsia'
%     'FF','00','FF', 'Magenta'
%     'BA','55','D3', 'MediumOrchid'
%     '93','70','DB', 'MediumPurple'
%     '99','66','CC', 'Amethyst'
%     '8A','2B','E2', 'BlueViolet'
%     '94','00','D3', 'DarkViolet'
%     '99','32','CC', 'DarkOrchid'
%     '8B','00','8B', 'DarkMagenta'
%     '6A','5A','CD', 'SlateBlue'
%     '48','3D','8B', 'DarkSlateBlue'
%     '7B','68','EE', 'MediumSlateBlue'
%     '4B','00','82', 'Indigo'
%     %Gray repeated with spelling grey
%     '80','80','80', 'Grey'
%     'D3','D3','D3', 'LightGrey'
%     'A9','A9','A9', 'DarkGrey'
%     '69','69','69', 'DimGrey'
%     '77','88','99', 'LightSlateGrey'
%     '70','80','90', 'SlateGrey'
%     '2F','4F','4F', 'DarkSlateGrey'


%% toPPT accepts the following predefined TeX parameters:
%     'alpha','upsilon','sim','beta','phi','leq','gamma','chi', ...
%     'infty','delta','psi','clubsuit','epsilon','omega','diamondsuit', ...
%     'zeta','Gamma','heartsuit','eta','Delta','spadesuit','theta', ...
%     'Theta','leftrightarrow','vartheta','Lambda','leftarrow','iota', ...
%     'Xi','uparrow','kappa','Pi','rightarrow','lambda','Sigma', ...
%     'downarrow','mu','Upsilon','circ','nu','Phi','pm','xi','Psi','geq', ...
%     'pi','Omega','propto','rho','forall','partial','sigma','exists', ...
%     'bullet','varsigma','ni','div','tau','cong','neq','equiv','approx', ...
%     'aleph','Im','Re','wp','otimes','oplus','oslash','cap','cup', ...
%     'supseteq','supset','subseteq','subset','int','in','o','rfloor', ...
%     'lceil','nabla','lfloor','cdot','ldots','perp','neg','prime', ...
%     'wedge','times','0','rceil','surd','mid','vee','varpi','copyright', ...
%     'langle','rangle','^','_','{','}','sqrt'



