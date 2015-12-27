function addBitmapGraphic(myFigure,myArg)

%% Powerpoint connecting

% Connect to PowerPoint
ppt = actxserver('PowerPoint.Application');

% Open current presentation
if get(ppt.Presentations,'Count') == 0
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





% % Identify current slide
% try
%     currSlide = wind.Selection.SlideRange.SlideNumber;
% catch %#ok<CTCH>
%     % No slides
% end
% 
% % Select the slide to which the figure will be exported
% slide_count = int32(get(op.Slides,'Count'));
% if strcmpi(myArg.defaultSlideNumber,'append')
%     slide = invoke(op.Slides,'Add',slide_count+1,11);
%     shapes = get(slide,'Shapes');
%     invoke(slide,'Select');
%     invoke(shapes.Range,'Delete');
% else
%     if strcmpi(myArg.defaultSlideNumber,'last')
%         slideNum = slide_count;
%     elseif strcmpi(myArg.defaultSlideNumber,'current');
%         slideNum = get(wind.Selection.SlideRange,'SlideNumber');
%     else
%         slideNum = myArg.defaultSlideNumber;
%     end
%     slide = op.Slides.Item(slideNum);
%     invoke(slide,'Select');
% end

[slide,~] = translateSlideNumberInformation2Slide(myArg,ppt,op,1);


%% Saveing the picture temporally

exportMode = myArg.defaultExportMode;%'matlab:emf';

exportFormatType = myArg.defaultExportFormatType;


imagePath = '';

exportModeDefined = 0;

if strcmp(exportMode,'exportFig')
    if strcmp(myArg.externalParameters,'')
        try
            export_fig(myFigure,[pwd '\pptfig.tif'],'-transparent',['-m',num2str(myArg.defaultMagnify)]);
            imageinfo = imfinfo([pwd '\pptfig.tif']);
            imagePath = [pwd '\pptfig.tif'];
        catch
            export_fig(myFigure,'.\pptfig.tif','-transparent',['-m',num2str(myArg.defaultMagnify)]);
            imageinfo = imfinfo('.\pptfig.tif');
            imagePath = '.\pptfig.tif';
        end

    else
        try
            export_fig(myFigure,[pwd '\pptfig.tif'],'-transparent',['-m',num2str(myArg.defaultMagnify)],myArg.externalParameters{1:length(myArg.externalParameters)});
            imageinfo = imfinfo([pwd '\pptfig.tif']);
            imagePath = [pwd '\pptfig.tif'];
        catch
            export_fig(myFigure,'.\pptfig.tif','-transparent',['-m',num2str(myArg.defaultMagnify)],myArg.externalParameters{1:length(myArg.externalParameters)});
            imageinfo = imfinfo('.\pptfig.tif');
            imagePath = '.\pptfig.tif';
        end

    end
    
    myArg.objectHeight = imageinfo.Height;
    myArg.objectWidth  = imageinfo.Width;
    exportModeDefined = 1;
end

doClipboardExport = 0;

if strcmp(exportMode,'matlab')
    
        
        
        [isAllowedFormat,indexAllowedFormat] = ismember(exportFormatType,myArg.allowedExportFormatType);
        
        if strcmp(exportFormatType,'clipboard')
            doClipboardExport = 1;
            exportModeDefined = 1;
        end

        if isAllowedFormat && ~doClipboardExport
            
            choosenExportFormat = myArg.allowedExportFormatType{indexAllowedFormat};
            choosenExportExt    = myArg.allowedExportFormatTypeExt{indexAllowedFormat};
        
            pathImage = fullfile(pwd, ['pptfig',choosenExportExt]);
            print(myFigure, ['-d',choosenExportFormat], pathImage); 

            imagePath = pathImage;

            pixelPos = getpixelposition(myFigure);

            myArg.objectHeight = pixelPos(4);
            myArg.objectWidth  = pixelPos(3);

            exportModeDefined = 1;
        
        end
        
        if ~exportModeDefined
            warning('No valid exportFormatType defined.');
        end
        
end

if ~exportModeDefined
    warning('No valid exportMode defined: Use "matlab" or "exportFig".');
end
    

if doClipboardExport
    
     pixelPos = getpixelposition(myFigure);
     hgexport(myFigure,'-clipboard');

     myArg.objectHeight = pixelPos(4);
     myArg.objectWidth  = pixelPos(3);
     
end

    
% end

% if doClip
%     
%     pixelPos = getpixelposition(myFigure);
%     hgexport(myFigure,'-clipboard');
%     
%     myArg.objectHeight = pixelPos(4);
%     myArg.objectWidth  = pixelPos(3);
%     
% end

%% Calculate position parameters%%%

postioningParameters = getPosParameters(myArg);

width  = postioningParameters.width;
height = postioningParameters.height;
top    = postioningParameters.top;
left   = postioningParameters.left;


%% Invoke into powerpoint
if ~doClipboardExport
    img = invoke(slide.Shapes,'AddPicture',imagePath,'msoFalse','msoTrue',left,top,width,height);
    invoke(img,'ZOrder','msoSendToBack');
    
    %% Clean up imagepath
    if myArg.cleanUpTempImages
        try
            delete(imagePath);
        catch
            warning('Cleaning up temp data failed.');
        end
    end
    
else
    img = invoke(slide.Shapes,'PasteSpecial');
    lastIndex = slide.Shapes.Range.Count;
    
    if lastIndex>10000
        slide.Shapes.Range(lastIndex).Top    = top;
        slide.Shapes.Range(lastIndex).Width  = width;
        slide.Shapes.Range(lastIndex).Height = height;
        slide.Shapes.Range(lastIndex).Left   = left;
    elseif lastIndex > 0;
        slide.Shapes.Range.Top    = top;
        slide.Shapes.Range.Width  = width;
        slide.Shapes.Range.Height = height;
        slide.Shapes.Range.Left   = left;
    end
end
% if doClip
%     
% %     imgShape = invoke(slide.Shapes,'AddShape','msoShapeRectangle',...
% %     left,top,width,height);
% 
%     img = invoke(slide.Shapes,'PasteSpecial');% DataType:=3 ' 3 = ppPasteMetafilePicture
%     img.Width = width;
%     
%     set(img,'Width',width);
% end

% for ii=1:slide.Shapes.Count
%     
%     ppShp = slide.Shapes.Shape(ii);
%     
%     try
%              ppShp.Top = 100;
%              ppShp.Left = 5;
%              ppShp.Height = 100;
%              ppShp.Width = 100;
%              display(num2str(ii));
%     catch
%         display('error');
%     end
% 
% end




end