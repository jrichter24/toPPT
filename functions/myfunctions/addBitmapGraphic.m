function addBitmapGraphic(myFigure,myArg)

%% Powerpoint connecting

% Connect to PowerPoint
ppt = actxserver('PowerPoint.Application');

% Open current presentation
if get(ppt.Presentations,'Count')==0
    op = invoke(ppt.Presentations,'Add');
else
    op = get(ppt,'ActivePresentation');
end

% Get information about the presentation


myArg.myPresentationHeight = op.PageSetup.SlideHeight; %in px
myArg.myPresentationWidth  = op.PageSetup.SlideWidth; %in px


% Set slide object to be the active pane
wind = get(ppt,'ActiveWindow');
panes = get(wind,'Panes');
slide_pane = invoke(panes,'Item',2);
invoke(slide_pane,'Activate');





% Identify current slide
try
    currSlide = wind.Selection.SlideRange.SlideNumber;
catch %#ok<CTCH>
    % No slides
end

% Select the slide to which the figure will be exported
slide_count = int32(get(op.Slides,'Count'));
if strcmpi(myArg.defaultSlideNumber,'append')
    slide = invoke(op.Slides,'Add',slide_count+1,11);
    shapes = get(slide,'Shapes');
    invoke(slide,'Select');
    invoke(shapes.Range,'Delete');
else
    if strcmpi(myArg.defaultSlideNumber,'last')
        slideNum = slide_count;
    elseif strcmpi(myArg.defaultSlideNumber,'current');
        slideNum = get(wind.Selection.SlideRange,'SlideNumber');
    else
        slideNum = myArg.defaultSlideNumber;
    end
    slide = op.Slides.Item(slideNum);
    invoke(slide,'Select');
end








%% Saveing the picture temporally

if strcmp(myArg.externalParameters,'')
    try
        export_fig(myFigure,[pwd '\pptfig.tif'],'-transparent',['-m',num2str(myArg.defaultMagnify)]);
        imageinfo = imfinfo([pwd '\pptfig.tif']);
    catch
        export_fig(myFigure,'.\pptfig.tif','-transparent',['-m',num2str(myArg.defaultMagnify)]);
        imageinfo = imfinfo('.\pptfig.tif');
    end
            
else
    try
        export_fig(myFigure,[pwd '\pptfig.tif'],'-transparent',['-m',num2str(myArg.defaultMagnify)],myArg.externalParameters{1:length(myArg.externalParameters)});
        imageinfo = imfinfo([pwd '\pptfig.tif']);
    catch
        export_fig(myFigure,'.\pptfig.tif','-transparent',['-m',num2str(myArg.defaultMagnify)],myArg.externalParameters{1:length(myArg.externalParameters)});
        imageinfo = imfinfo('.\pptfig.tif');
    end

end


%% Calculate position parameters%%%

myArg.objectHeight = imageinfo.Height;
myArg.objectWidth  = imageinfo.Width;

postioningParameters = getPosParameters(myArg);

width  = postioningParameters.width;
height = postioningParameters.height;
top    = postioningParameters.top;
left   = postioningParameters.left;


%% Invoke into powerpoint
img = invoke(slide.Shapes,'AddPicture',[pwd '\pptfig.tif'],...
    'msoFalse','msoTrue',left,top,width,height);

invoke(img,'ZOrder','msoSendToBack');

end