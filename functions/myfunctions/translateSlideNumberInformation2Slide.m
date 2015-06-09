function [slide,slideNum] = translateSlideNumberInformation2Slide(myArg,ppt,op,doManipulateOp)


    %doManipulateOp => do we want to add slides or do we just want to get
    %information about the selcted SlideNumber?

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
        % No slides - change  current to append if necessary
        if(strcmpi(myArg.defaultSlideNumber,'current'))
            myArg.defaultSlideNumber = 'append';
        end
    end

    % Select the slide to which the figure will be exported
    slide_count = int32(get(op.Slides,'Count'));
    
    blockIncrease = 0;
    if strcmpi(myArg.defaultSlideNumber,'append')
        
        slideNum = slide_count+1;
        
        if doManipulateOp
            slide = invoke(op.Slides,'Add',slideNum,11);
            shapes = get(slide,'Shapes');
            invoke(slide,'Select');
            invoke(shapes.Range,'Delete');
        else
            % We overwrite the slide 
            slide = [];
        end
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
            
            if strcmpi(myArg.defaultSlideAddMethod,'insert') || strcmpi(myArg.defaultSlideAddMethod,'after')
                slideNum = myArg.defaultSlideNumber+1; % +1 because we want to add AFTER the desired title slide
                blockIncrease = 1;
            elseif strcmpi(myArg.defaultSlideAddMethod,'update') || strcmpi(myArg.defaultSlideAddMethod,'before')
                slideNum = myArg.defaultSlideNumber;
            else
                slideNum = myArg.defaultSlideNumber;
            end
            
        else
            slideNum = myArg.defaultSlideNumber;
        end
        
        
        %% This will add empty slides in case the user wants to add content to none existing slideNumbers
        if isnumeric(slideNum)
            
            
            if strcmpi(myArg.defaultSlideAddMethod,'insert') || strcmpi(myArg.defaultSlideAddMethod,'after')
                if ~blockIncrease
                    slideNum = slideNum+1; % +1 because we want to add AFTER the desired title slide
                end
            elseif strcmpi(myArg.defaultSlideAddMethod,'update') || strcmpi(myArg.defaultSlideAddMethod,'before')
                %slideNum = myArg.defaultSlideNumber;
            else
                %slideNum = myArg.defaultSlideNumber;
            end
            
            if doManipulateOp
                
                if slide_count<slideNum %% There are not enough slides => add dummy slides
                    for ii=1:(slideNum-slide_count)
                        slide = invoke(op.Slides,'Add',slide_count+ii,11);
                        shapes = get(slide,'Shapes');
                        invoke(slide,'Select');
                        invoke(shapes.Range,'Delete');
                    end
                else 

                % Insert new slide - content will be insert later

                if strcmpi(myArg.defaultSlideAddMethod,'insert') && doManipulateOp

                    slide = invoke(op.Slides,'Add',slideNum,11);
                    shapes = get(slide,'Shapes');
                    invoke(slide,'Select');
                    invoke(shapes.Range,'Delete'); 

                end

                % Update content to existing slide  
                % Do nothing here   
            
                end
                
            end
        
        

        end
        
        try
            slide = op.Slides.Item(slideNum);
            invoke(slide,'Select');
        catch
            slide = [];
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