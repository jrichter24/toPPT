function addText(myArg)
% Help ADDTEXT -  CSS Focused:
%
% Css-Tags (basic):
% ---------------
%   When adding text with toPPT the following basic css-Tags are accepted:
%
%   Example:
%   <b>Your bold Text</b>
%   <i>Your italic Text</i>
%   <u>Your underlined Text</u> 
%
%   Tags can also be combined:
%   <u><i><b>Your bold, italic and underlined Text</b></i></u>
%   
%   It is absolutely necessary to always close open tags (e.g. opening tag <u> 
%   and closing tag </u>).
%
% 
% Css-Tags (special):
% ---------------
%   In addition there is a special tag <s>. This tag can be used to change
%   the color, size and the font of the text:
%   
%   Example:
%   <s color:red>Red text</s> - => Use 'help rgb' for a list of all
%       accepted colors
%
%   <s color:#5F9EA0> Hex Color Text</s> - A color can be defined as
%       Hex-String. It is important to use a '#' before the color code.
%
%   <s font-family:Times New Roman> Times New Roman Text </s>
%   <s font-size:40>BIG TEXT</s>
% 
%   Tags can also be combined:
%   <s color:#5F9EA0; font-size:22; font-family:Aharoni> Mixed text</s>
%
%   It is important to separate all parameters in the <s> tag with ';' and
%   to close the opening tag <s> with a closing </s> tag.
%
%
% Css-Tags (basic and special mixed):
% ---------------
%   Basic and speical tags can also be mixed:
%
%   Example:
%   <b><s color:blue; font-size:22; font-family:Aharoni>Bold blue big Aharoni text</s></b>'

%% The argument probably just contains a string that should be added as a bulletpoint


if length(myArg.externalParameters) >= 1
    
    myFormattedText = toLinebreakVersion(myArg.externalParameters{1});
    if myFormattedText == -1
        error('Invalid input');
    else
        myArg.doSetText = 1;
    end
    
end

%% Powerpoint connecting

% Connect to PowerPoint


ppt = actxserver('PowerPoint.Application');


% Open current presentation
if get(ppt.Presentations,'Count')==0
    op = invoke(ppt.Presentations,'Add');
else
    op = get(ppt,'ActivePresentation');
end




%% If we want to addText tables etc eq. myArg.isGeneralCommand == 0 => go on
if ~myArg.isGeneralCommand

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
        
        %In case the slideNumber
        
    catch
        % No slides - change append to current if necessary
        if(strcmpi(myArg.defaultSlideNumber,'current'))
            myArg.defaultSlideNumber = 'append';
        end
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
            try
                slideNum = get(wind.Selection.SlideRange,'SlideNumber');
            catch MyError
                error('You have to add at least one slide first, or simply set SlideNumber to append in the call')
            end
        elseif ischar(myArg.defaultSlideNumber) 
            % We want to add the content "near" a slideTitle for this we
            % have to get all SlideTitles and compare            
            
            slideNumberIndex = getSlideNumberByTitle(myArg,op,slide_count);
            
            myArg.defaultSlideNumber = slideNumberIndex;
            
            if strcmpi(myArg.defaultSlideAddMethod,'insert')
                slideNum = myArg.defaultSlideNumber+1; % +1 because we want to add AFTER the desired title slide
            else
                slideNum = myArg.defaultSlideNumber;
            end
            
        else
            slideNum = myArg.defaultSlideNumber;
        end
        
        
        %% This will add empty slides in case the user wants to add content to an non existing slideNumber
        if isnumeric(slideNum)
            if slide_count<slideNum %% There are not enough slides => add dummy slides
                for ii=1:(slideNum-slide_count)
                    slide = invoke(op.Slides,'Add',slide_count+ii,11);
                    shapes = get(slide,'Shapes');
                    invoke(slide,'Select');
                    invoke(shapes.Range,'Delete');
                end
            else 

            % Insert new slide - content will be insert later
            
            if strcmpi(myArg.defaultSlideAddMethod,'insert')
                
                slide = invoke(op.Slides,'Add',slideNum,11);
                shapes = get(slide,'Shapes');
                invoke(slide,'Select');
                invoke(shapes.Range,'Delete'); 
                
            end
        
            % Update content to existing slide  
            % Do nothing here   
                
            end
        
        

        end
        
        slide = op.Slides.Item(slideNum);
        invoke(slide,'Select');
    end



    %% Add some additional values


    myArg.objectHeight = 100;
    myArg.objectWidth  = 100;

    postioningParameters = getPosParameters(myArg);

    myArg.userWidth  = postioningParameters.width;
    myArg.userHeight = postioningParameters.height;
    myArg.userTop    = postioningParameters.top;
    myArg.userLeft   = postioningParameters.left;


    %% Different cases for adding text to slide

    %Case: Set Title
    if myArg.doSetTitle
       setTitle(myArg,slide) 
    end

    %Case: Set Text
    if myArg.doSetText
        setText(myArg,slide,myFormattedText)   
    end

    %Case: Set Table
    if myArg.doSetTable
        setTable(myArg,slide)   
    end

    %Case: Set HyperLink
    if myArg.doSetHyper
        setHyperLink(myArg,slide,ppt);
    end

else %% We want to perform a gerneal Command like saving
       
    if myArg.doApplyTemplate
        applyTemplate(myArg,op);
    end
    

    
    %% Saving presentation
    if myArg.doSavePPTPath || myArg.doSavePPTFilename
        savePPT(myArg,op);
    end
    
    %% Close presentation
    if myArg.doClose
        closePPT(myArg,op)
    end

    
end



end


%% Get slideNumber by Title
function slideNumberIndex = getSlideNumberByTitle(myArg,op,slide_count)

    titleStruct = cell(1,slide_count);
    editDistanceArray = zeros(1,slide_count);

    for ii=1:slide_count

        try
            titleStruct{ii} = get(op.Slides.Item(ii).Shapes.Title.TextFrame.TextRange,'Text');
        catch % No title present
            titleStruct{ii} = num2str(ii);
        end

        if numel(titleStruct{ii}) < numel(myArg.defaultSlideNumber)
            editDistanceArray(ii) = EditDistance(titleStruct{ii},myArg.defaultSlideNumber);
        else
            editDistanceArray(ii) = EditDistance(titleStruct{ii},myArg.defaultSlideNumber(1:numel(myArg.defaultSlideNumber)));
        end
        %% If EditDistance is zero we are already done
        if editDistanceArray(ii) == 0
            break;
        end

    end

    %% Identify slide
    [~,slideNumberIndex] = min(editDistanceArray);
end

%% Apply template
function closePPT(myArg,presentation)
    try
        presentation.Saved = 1;
        presentation.Close;
        
    catch
        display('Closing presentation failed!');
    end
end


%% Apply template
function applyTemplate(myArg,presentation)
    try
        presentation.ApplyTemplate(myArg.templatePath);
    catch
        display(['Applying template failed! - Right templatepath? Path: ',myArg.templatePath]);
    end
end

%% Save ppt functions

function savePPT(myArg,presentation)

    defaultExtension = 'pptx';

    if(myArg.doSavePPTPath && myArg.doSavePPTFilename)
        
        fullpath = [myArg.savePath,'./',myArg.saveFilename,'.',defaultExtension];
        %% Save
        try
        presentation.SaveAs(fullpath);
        catch
            display('Saving presentation failed!');
        end
        
    elseif (myArg.doSavePPTPath)
        %% We have to autogenerate a filename
        filename = [strrep(datestr(now),':',''),'_new_presentation'];
        fullpath = [myArg.savePath,'\',filename,'.',defaultExtension];
        %% Save
        try
        presentation.SaveAs(fullpath);
        catch
            display('Saving presentation failed!');
        end
        
    elseif (myArg.doSavePPTFilename)
        %% We are saving the presentation to the active directory
        fullpath = [pwd,'.\',myArg.saveFilename,'.',defaultExtension];
        %% Save
        try
        presentation.SaveAs(fullpath);
        catch
            display('Saving presentation failed!');
        end
    end

    
    
end

%%
function setTitle(myArg,slide)

    myArg.defaultTitle = toLinebreakVersion(myArg.defaultTitle);
    
    try
        if(myArg.doTexText)
            
            try
                tex2ppt(slide.Shapes.Title.TextFrame.TextRange,myArg.defaultTitle,0); %% update title
            catch
                tex2ppt(slide.Shapes.AddTitle.TextFrame.TextRange,myArg.defaultTitle,0); %% add title
            end
            
        else
            try
                set(slide.Shapes.Title.TextFrame.TextRange,'Text',myArg.defaultTitle);
            catch
                set(slide.Shapes.AddTitle.TextFrame.TextRange,'Text',myArg.defaultTitle);
            end
        end
    catch
        display('Error using Tex interpreter - please check if your tex code has no errors')
        try
               set(slide.Shapes.Title.TextFrame.TextRange,'Text',myArg.defaultTitle);
        catch
               set(slide.Shapes.AddTitle.TextFrame.TextRange,'Text',myArg.defaultTitle);
        end
    end
    
    
%     try
%         set(slide.Shapes.Title.TextFrame.TextRange,'Text',myArg.defaultTitle);
% 
%     catch myError
% 
%         set(slide.Shapes.AddTitle.TextFrame.TextRange,'Text',myArg.defaultTitle);
% 
%     end
    
end

%%
function setText(myArg,slide,myFormattedText)

    myNewTextBox = invoke(slide.Shapes,'AddTextbox','msoTextOrientationHorizontal',myArg.userLeft,myArg.userTop,myArg.userWidth,myArg.userHeight);%Working
   
    % First lets see if parts of the text has to be bold or italic or
    % underlined
    %set(myNewTextBox.TextFrame.TextRange,'Text',myFormattedText); %Working
    
    intperetHtml(myFormattedText,myNewTextBox.TextFrame.TextRange,myArg);
    
    
    %Adding bullets - can be truned off if user does not want that
    
    if myArg.addBulletPoints
    %ItemsBulletedList
       
        if myArg.addBulletNumbers
            set(myNewTextBox.TextFrame.TextRange.ParagraphFormat.Bullet,'Visible',1);%Working
            set(myNewTextBox.TextFrame.TextRange.ParagraphFormat.Bullet,'RelativeSize',1);%Working
            set(myNewTextBox.TextFrame.TextRange.ParagraphFormat.Bullet.Font,'Name','Symbol');%Working
            %set(myNewTextBox.TextFrame.TextRange.ParagraphFormat.Bullet,'Style','ppBulletCircleNumWDBlackPlain');%WorkingCircle
        else
            set(myNewTextBox.TextFrame.TextRange.ParagraphFormat.Bullet,'Visible',1);%Working
            set(myNewTextBox.TextFrame.TextRange.ParagraphFormat.Bullet,'RelativeSize',1);%Working
            set(myNewTextBox.TextFrame.TextRange.ParagraphFormat.Bullet,'Type','ppBulletUnnumbered');%WorkingCircle
        end
 
        
    end
end



function myFormattedText = intperetHtml(myFormattedText,textRange,myArg)

    hasError = 0;

    % knownCommands ALLWAYS have to have the length 3(opening tag) or 4 (closing tag)
    knownCommands = {{'<b>','</b>','<\b>'},{'<u>','</u>','<\u>'},{'<i>','</i>','<\i>'},{'<s>','</s>','<\s>'}};
    
    isSpecialCommand = [0,0,0,1]; % Does the knowCommand has some special attributes?
    knownSpecialCommandsAttributes = {{'s','color:','font-family:','font-size:'}}; % The first element of the structure is the known tag itsself without brackets

    bStart = cell(1,5);
    bStop = cell(1,5);
    
    bStartSpecialCommand = cell(1,5);
    bStopSpecialCommand  = cell(1,5);
    
    % First thing we have to do is finding the special Commands try to get
    % the attributes
    specialCommands = knownCommands(logical(isSpecialCommand));
    
    for jj=1:numel(specialCommands)
        myFormattedText = strrep(myFormattedText,specialCommands{jj}{3},specialCommands{jj}{2}); % Helps if for some reason a backslash was used
        % Now we have to create an object that holds all attributes we
        % found - we only search for a part of the opening tag e.g. <s
        % instead of <s>
        bStartSpecialCommand{jj} = strfind(myFormattedText,specialCommands{jj}{1}(1:end-1)); % This hold all starting points of special commands
        bStopSpecialCommand{jj} = strfind(myFormattedText,specialCommands{jj}{2});
    end
    
    % In the next phase we are going to extract the attributes => saving
    % them to an object in the order of apperiance => deleting them from
    % text
    
    
    offsetSpecialCommand = 0;
    
    for jj = 1:numel(bStartSpecialCommand)
        for ii = 1:numel(bStartSpecialCommand{jj})
            
            currentStart = bStartSpecialCommand{jj}(ii)-offsetSpecialCommand;
            currentEnd   = bStopSpecialCommand{jj}(ii)-offsetSpecialCommand;
            % Now we save the whole attributestring in to an object
            currentUnformatedAttributeString = myFormattedText(currentStart:currentEnd-1); % this includes e.g.
            currentSpecialCommand = myFormattedText(currentStart+1:currentStart+1);
            % Now we use a helper function that is able to form a strucutre
            % with the current attributes
            attributeCell{jj}{ii} = getAttributesFromString(currentUnformatedAttributeString,currentSpecialCommand,knownSpecialCommandsAttributes);
            
            % Now we can overwrite <s ......> with <s> and assign the
            % attributes later
            closeTagIndex = strfind(currentUnformatedAttributeString,'>');
            myFormattedText = [myFormattedText(1:currentStart-1),'<',currentSpecialCommand,'>',myFormattedText(currentStart+closeTagIndex:end)];
            
            % Because we deleted some of the text we have to assign this as
            % offset for the next current Start and End
            offsetSpecialCommand = offsetSpecialCommand+closeTagIndex-3;
        end
    end
    
    %The variable attributeCell{1, 1}{1, 3} does not exist.

    % Find tags <b>,</b>, <u>,</u>, <i>, </i>
    
    % Helps if for some reason a backslash was used
    for jj=1:numel(knownCommands)
        myFormattedText = strrep(myFormattedText,knownCommands{jj}{3},knownCommands{jj}{2});
    end
    
    % We have to get a previev of what the tex code will do - otherwise
    % later on all positions will be shifted after applying the
    % texConverter
    if(myArg.doTexText)
        try
            PreviewMyFormattedText = tex2ppt(0,myFormattedText,1);
        catch
            PreviewMyFormattedText = myFormattedText;
        end
    else
        PreviewMyFormattedText = myFormattedText;
    end

    
    for jj=1:numel(knownCommands)

        bStart{jj} = strfind(PreviewMyFormattedText,knownCommands{jj}{1});
        bStop{jj}  = strfind(PreviewMyFormattedText,knownCommands{jj}{2});
        
    end
    
    
    myAddText = myFormattedText;
    
    vOrgIndices = 1:length(myFormattedText);
    vNewIndices = vOrgIndices;
    
    

    % Second create indice map and delete all tags
     for jj=1:numel(knownCommands)
        for kk=1:length(myFormattedText)
            if find(bStart{jj} == kk)> 0
                vNewIndices = [vNewIndices(1:kk),ones(1,length(knownCommands{jj}{1}))*vNewIndices(kk),vNewIndices(kk+length(knownCommands{jj}{1})+1:end)-length(knownCommands{jj}{1})];
                
                % Do the same thing for the closing tags
                pp = bStop{jj}(bStart{jj} == kk);
                vNewIndices = [vNewIndices(1:pp),ones(1,length(knownCommands{jj}{2}))*vNewIndices(pp),vNewIndices(pp+length(knownCommands{jj}{2})+1:end)-length(knownCommands{jj}{2})];
            end
        end
        
        myAddText = strrep(myAddText,knownCommands{jj}{1},'');
        myAddText = strrep(myAddText,knownCommands{jj}{2},'');
        
     end
    

    %% In case Tex was used we run tex2ppt
    
%     set(textRange,'Text',myAddText); %No Tex

    %% Do the tex interpretation if possible
    try
        if(myArg.doTexText)
            tex2ppt(textRange,myAddText,0);
        else
            set(textRange,'Text',myAddText); %No Tex
        end
    catch
         display('Error using Tex interpreter - please check if your tex code has no errors')
         set(textRange,'Text',myAddText);
    end

    selementindex = 0;
    
    % Now assign all properties
    for jj=1:numel(knownCommands) 

        % Is the number of opening tags and closing tags the same?
        if numel(bStart{jj}) == numel(bStop{jj}) && numel(bStop{jj})>=1
            for ii=1:numel(bStart{jj})
                switch knownCommands{jj}{1}
                    case '<b>'
                        set(textRange.Characters(vNewIndices(bStart{jj}(ii)), vNewIndices(bStop{jj}(ii))-(vNewIndices(bStart{jj}(ii)))).Font,'Bold',1);
                    case '<u>'
                        set(textRange.Characters(vNewIndices(bStart{jj}(ii)), vNewIndices(bStop{jj}(ii))-(vNewIndices(bStart{jj}(ii)))).Font,'Underline',1);
                    case '<i>'
                        set(textRange.Characters(vNewIndices(bStart{jj}(ii)), vNewIndices(bStop{jj}(ii))-(vNewIndices(bStart{jj}(ii)))).Font,'Italic',1);
                    case '<s>'
                        selementindex = selementindex +1;
                        attributeSpecialCommandIndex = 1;
                        rangeStart = vNewIndices(bStart{jj}(ii));
                        rangeStop  = vNewIndices(bStop{jj}(ii))-(vNewIndices(bStart{jj}(ii)));
                        
                        assignStyleToTextElement(textRange,rangeStart,rangeStop,attributeCell{attributeSpecialCommandIndex}{selementindex})
                        %set(textRange.Characters(vNewIndices(bStart{jj}(ii)), vNewIndices(bStop{jj}(ii))-(vNewIndices(bStart{jj}(ii)))).Font.Color,'RGB',16777215); % 16777215 = 256^3-1
                end
                
            end

        else
            hasError = 1;
        end
        
       
    
    end
    
    %% Is text a hyperlink?
    
    % Test 
    isHyperlink = 1;
    
    
% For Each osld In ActivePresentation.Slides
% For Each oshp In osld.Shapes
% If oshp.HasTextFrame Then
% If oshp.TextFrame.TextRange = "test" Then
% With oshp.TextFrame.TextRange.ActionSettings(ppMouseClick)
% .Action = ppActionHyperlink
% .Hyperlink.Address = ""
% .Hyperlink.SubAddress = strID & "," & strIndx & "," & strTitle
% End With
% End If
% End If
% End If
% Next oshp
% Next osld
% End Sub
    
%     if isHyperlink
%         set(textRange.ActionSettings(ppMouseClick).Action, ' ppActionHyperlink');
%         set(textRange.ActionSettings.PpMouseActivation.ppMouseClick);
%     end
    
    % Test end
    

    

end

function assignStyleToTextElement(textRange,rangeStart,rangeStop,currentAttributeCell)

    % First have to see what we will have to do
    for ii=2:numel(currentAttributeCell)
        %% Get current Attribute
        
        if ~isempty(currentAttributeCell{ii})
        currentAttribute = currentAttributeCell{ii}{1}{1};
        currentValue = currentAttributeCell{ii}{2}{1};
        
            switch currentAttribute
                case 'color:'
                    % the value for color can be a hexnumber or a stringvalue
                    % (like red, blue etc.) but we will have to transform it to
                    % a dec number for ppt
                    try
                        set(textRange.Characters(rangeStart,rangeStop).Font.Color,'RGB',rgb(currentValue,'ppt'));
                    catch
                        error(['An error occured when setting the Font-Color: ',currentValue,' only names are allowed (red, etc.) and hex values in the form #hexnumber']);
                    end
                case 'font-family:'
                    try
                        set(textRange.Characters(rangeStart,rangeStop).Font,'Name',currentValue);
                    catch
                        error(['An error occured when setting the Font-Family: ',currentValue]);
                    end
                case 'font-size:'
                    try
                        set(textRange.Characters(rangeStart,rangeStop).Font,'Size',str2num(currentValue));
                    catch
                        error(['An error occured when setting the Font-Family: ',currentValue,' / you just have to set a number no px or % is allowed']);
                    end
            end
        end
        
    end

end


function attributeCell = getAttributesFromString(currentUnformatedAttributeString,currentSpecialCommand,knownSpecialCommandsAttributes)
    
    % First of all we have to find the desired set of
    % SpecialCommandsAttributes for the current currentSpecialCommand
    for ii = 1:numel(knownSpecialCommandsAttributes)
        if(strcmp(knownSpecialCommandsAttributes{ii}{1},currentSpecialCommand))
            % Found
            currentSpecialCommandSet = knownSpecialCommandsAttributes{ii};
        end
    end
    
    % Now we have to find the attributes from the currentSpecialCommandSet
    % and the values for this attribute -  attributes can be separated by
    % spaces or if it is the last attribute possibly their is a >
    for ii = 2:numel(currentSpecialCommandSet)
        curStartSingleSpecialCommand = strfind(currentUnformatedAttributeString,currentSpecialCommandSet{ii});
        
        if~isempty(curStartSingleSpecialCommand)
            curAttributeValue = getSubstring(curStartSingleSpecialCommand,currentUnformatedAttributeString,':',{'>',';'});

            % Persist to structure;
            attributeCell{ii} = {{currentSpecialCommandSet{ii}},{strrep(curAttributeValue,'"','')}};% We delete all " in value if they are their
        end
        
    end

end

function substring = getSubstring(offset,fullString,charStart,charStop)

    captureMe = 0;
    stopMe    = 0;
    substring = '';

    for ii = offset:numel(fullString)
        currentChar = fullString(ii);

        if (captureMe)
 
            if iscell(charStop)
                for jj=1:numel(charStop)
                    if strcmp(currentChar,charStop{jj})
                       stopMe = 1;
                       break;
                    end
                end 

            elseif ischar(charStop)
                if strcmp(currentChar,charStop)
                    stopMe = 1;
                    break;
                end
            end
            
            if(stopMe)
                break;
            end
            
            
            substring = [substring,currentChar];
        end
        
        %Second lets find the start char
        if strcmp(currentChar,charStart)
            % Start found! Now we are capturing the substring
            captureMe = 1;
        end
        
    end
end

%%
function setTable(myArg,slide)
%     
%     mytable = slide.Shapes.AddTable(3, 4, 10, 10, 288, 216);
%     set(mytable.Table.Cell(1, 1).Shape.TextFrame.TextRange,'Text','Test11');
%     set(mytable.Table.Cell(1, 2).Shape.TextFrame.TextRange,'Text','Test12');
    
    %Add table
    pptTableObject = toPPTTable(myArg);
    myNewTable = slide.Shapes.AddTable(pptTableObject.rowCount,pptTableObject.colCount,myArg.userLeft,myArg.userTop,myArg.userWidth,myArg.userHeight);
    
    %Insert Data into Table
    
    for ii=1:pptTableObject.rowCount
        for jj=1:pptTableObject.colCount
            set(myNewTable.Table.Cell(ii, jj).Shape.TextFrame.TextRange,'Text',pptTableObject.myCellTable{ii,jj});
        end
    end
    
end

%%
function setHyperLink(myArg,slide,ppt)

%% Todo

%     myNewTextBox   = invoke(slide.Shapes,'AddTextbox','msoTextOrientationHorizontal',myArg.userLeft,myArg.userTop,myArg.userWidth,myArg.userHeight);%Working
%     %myNewTextRange = myNewTextBox.TextFrame.TextRange;
%     myNewTextRange = myNewTextBox.TextFrame.TextRange;
%     
%     %oSlide.Shapes[1].ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Action = PowerPoint.PpActionType.ppActionRunProgram;
%     %set(myNewTextRange.ActionSettings(.PpMouseActivation.ppMouseClick));
%     %invoke(myNewTextRange,'ActionSettings','ppMouseClick');
%     
%     
%     
%     myNewTextRangeActionSettings = myNewTextRange.ActionSettings;
%     asetmc = set(myNewTextRangeActionSettings,'PpMouseActivation','ppMouseClick');
%     
%     set(myNewTextBox.TextFrame.TextRange.ParagraphFormat.Bullet,'Type','ppBulletUnnumbered');%WorkingCircle
%     
%     set(myNewTextRangeActionSettings,'Action','ppActionHyperlink');
%     
% % PowerPoint.ActionSetting asetmc =aset[PowerPoint.PpMouseActivation.ppMouseClick];
% % if (asetmc.Action == PowerPoint.PpActionType.ppActionHyperlink)
%     
%     
% % %     > PowerPoint.ActionSettings aset = tr.ActionSettings;
% % % > PowerPoint.ActionSetting asetmc =
% % % > aset[PowerPoint.PpMouseActivation.ppMouseClick];
% % % > if (asetmc.Action == PowerPoint.PpActionType.ppActionHyperlink)
%     
%     
%     
%     set(myNewTextRange.ActionSettings(ppt.PpMouseActivation.ppMouseClick),'Action','ppActionHyperlink');
%     
%     
%     
%     myNewTextBox.ActionSettings(ppMouseClick).Action = 'ppActionLastSlide';
%     set(myNewTextRange,'Text','test');
%     
%     
    set(myNewTextRange.ActionSettings,'Action','ppActionHyperlink');
%     
%     
%     set(myNewTextRange.ActionSettings.Hyperlink,'Address','first');
%     set(myNewTextRange.ActionSettings.Hyperlink,'SubAddress',3);
    
%     With ActivePresentation.Slides(1).Shapes(1).GroupItems(1).TextFrame.TextRange.ActionSettings(ppMouseClick)
%     .Action = ppActionHyperlink
%     '.Hyperlink.Address = SlideNumber
%     .Hyperlink.SubAddress = 3
% End With
%     
    
% With shp.TextFrame.TextRange.Characters(m.FirstIndex + 1, m.Length) 
%     With .ActionSettings(ppMouseClick) 
%         .Hyperlink.Address = fullURL 
%         .Hyperlink.TextToDisplay = m.Value 
%     End With 
% End With 
end

function myText = toLinebreakVersion(text)
 myText = '';

 isCell       = 0;
 isPureString = 0;
 %% Check if cellarray or string
 
 if ischar(text) %The user wants to add text to the slide
     isPureString = 1;
     myText = text;
 end
 
 %testtext = ['First Item', char(13), 'Second Item', char(13) , 'Third Item'];%Working
 
 if iscell(text) %The user wants to add text to the slide
     isCell = 1;
     
     for ii=1:length(text)
         if ii ~= length(text)
            myText = [myText,text{ii},char(13)];
         else
            myText = [myText,text{ii}];
         end
     end
 end
 
 %% If nor text nor cell return a -1
 
 if isCell == 0 && isPureString == 0
    myText = -1;
 end
 
 
end

function pptTableObject = toPPTTable(myArg)

% we have to see if the first element of myArg.defaultTable{1} is a cell of
% string if yes => the strings will be the title for our table, if it just
% a string we will only have one 


%Case strings, matrix
%For argument = Strings, Matrix of numbers
colCount = numel(myArg.defaultTable{1}); % This is the caption count


%% Now we have to distinguish betwenn different supported cases

% Case 1: Captions and as myArg.defaultTable{1} is a numerical table

if isnumeric(myArg.defaultTable{2})
    
    % Now we have to possibly rotate the matrix to fit the number of
    % collums
    
    if(size(myArg.defaultTable{2},2) ~= colCount)
        % Rotate it
        myArg.defaultTable{2} = myArg.defaultTable{2}';
    end
    
    % Check again
    if(size(myArg.defaultTable{2},2) ~= colCount)
        % Number of elements is not fitting
        error('Number of elements is not fitting to collums'); % TODO more information
    else
        
        % Create table
        rowCount = length(myArg.defaultTable{2});

        myCellTable = cell(rowCount+1,colCount);

        for jj=1:colCount

            myCellTable{1,jj} = num2str(myArg.defaultTable{1}{jj});
        end

        for ii=1:rowCount
            for jj=1:colCount
                myCellTable{ii+1,jj} = num2str(myArg.defaultTable{2}(ii,jj));
            end
        end
        
    end
 
    
% Case 2 defaultTable{2} is a cell
elseif iscell(myArg.defaultTable{2})
    
    isColVectors = 1; % By default we assume ColVectors
    isFullReadyCell   = 0; % Maybe the cell already include all necessary data - is multidimensional
    
    % Check if myArg.defaultTable{1} is multidimensional - if not it will
    % be treated as row or collumn vec
    if size(myArg.defaultTable{2}{1},1) >1  && size(myArg.defaultTable{2}{1},2) >1 && numel(myArg.defaultTable{2}) == 1
        
        % Lets check if one dimension fits the colCount
        if size(myArg.defaultTable{2}{1},1) ~= colCount
            % Rotate it
            myArg.defaultTable{2}{1} = myArg.defaultTable{2}{1}';
            rowCount = size(myArg.defaultTable{2}{1},2);
        else
            rowCount = size(myArg.defaultTable{2}{1},2);
        end
        
        % Check again
        if(size(myArg.defaultTable{2}{1},1) ~= colCount)
            % Number of elements is not fitting
            error('Number of elements is not fitting to collums'); % TODO more information
        else
            
            myCellTable = cell(rowCount+1,colCount);
            
            % Check if cell is allready strCell
            if iscellstr(myArg.defaultTable{2}{1})
                %myCellTable = myArg.defaultTable{2}{1};
                isFullReadyCell = 1;
                
            else
                myArg.defaultTable{2}{1} = any2Str(myArg.defaultTable{2}{1});
                isFullReadyCell = 1;
            end
            
            % Add header
            for jj=1:colCount
               myCellTable{1,jj} = num2str(myArg.defaultTable{1}{jj});
            end
            
             for ii=1:rowCount
                 for jj=1:colCount
                    myCellTable{ii+1,jj} = any2Str(myArg.defaultTable{2}{1}{jj,ii});       
                 end
             end
        
        end
        
    end
    
    
    if ~isFullReadyCell
        
    % Each element of myArg.defaultTable{2} can be a cell too. It also can
    % be a row or a collum vector -  to keep it easy we only allow the user
    % to either deliver only collum OR rowvectors
    
        if numel(myArg.defaultTable{2}{1}) ~= colCount
            % Possibly the user wants Rowvectors
            isColVectors = 0;
            rowCount = numel(myArg.defaultTable{2}{1});
        else
            isColVectors = 1;
            rowCount = numel(myArg.defaultTable{2}); % The number of elements is defining the number of rows
        end
        
            
        myCellTable = cell(rowCount+1,colCount);

        for jj=1:colCount
            myCellTable{1,jj} = num2str(myArg.defaultTable{1}{jj});
        end

        switch isColVectors
            case 0

                for ii=1:rowCount
                    for jj=1:colCount
                        if isnumeric(myArg.defaultTable{2}{jj})
                            myCellTable{ii+1,jj} = any2Str(myArg.defaultTable{2}{jj}(ii));
                        else
                            myCellTable{ii+1,jj} = any2Str(myArg.defaultTable{2}{jj}{ii});
                        end
                    end
                end

            case 1
                for ii=1:rowCount
                    for jj=1:colCount
                        if isnumeric(myArg.defaultTable{2}{ii})
                            myCellTable{ii+1,jj} = any2Str(myArg.defaultTable{2}{ii}(jj));
                        else
                            myCellTable{ii+1,jj} = any2Str(myArg.defaultTable{2}{ii}{jj});
                        end
                    end 
                end

        end
        
        if ~iscellstr(myCellTable)
            error('Wrong input data')
        end

        
    end
    
    
    
    
end



pptTableObject.rowCount    = rowCount+1;
pptTableObject.colCount    = colCount;
pptTableObject.myCellTable = myCellTable;

end

% THis function trasfroms any input to a string
function out = any2Str(in)

if isnumeric(in) && numel(in) == 1
    out = num2str(in);
elseif ischar(in)
    out = in;
elseif isnumeric(in) && numel(in) > 1
    % We first convert the matrix to a cell before we are also going to
    % create strings
    in = num2cell(in);
end

if iscell(in)
    out = cell(size(in));
    
    for ii=1:size(in,1)
        for jj=1:size(in,2)
            out{ii,jj} = any2Str(in{ii,jj});
        end
    end
    
end



end
