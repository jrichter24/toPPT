%function addVectorGraphic(myFigure,arguments)
function addVectorGraphic(myFigure,myArg)

%%%% Here is the start %%%

% 

H = myFigure;

B2 = copyfig(H);

H2  = myFigure;


%Further checking and error handling is done in pptfigure by Dimitri

%Second we have to check if the figure fig has a colorbar
hasColorbar = 1;

mycolorbar = findobj(get(H2,'Children'),'Tag','Colorbar');

if isempty(mycolorbar)
    hasColorbar = 0;
end

if hasColorbar == 1 %we have to split the figure into to inpendent figures

    ch=findall(H2,'tag','Colorbar');
    delete(ch);

    if strcmp(myArg.externalParameters,'')
        %pptfigure(H2,'AutoGroupMin',myArg.defaultAutoGroupMin,'SlideNumber',myArg.defaultSlideNumber);
        pptfigure(H2,myArg);
    else
        pptfigure(H2,myArg);
        %pptfigure(H2,'AutoGroupMin',myArg.defaultAutoGroupMin,'SlideNumber',myArg.defaultSlideNumber,myArg.externalParameters{1:length(myArg.externalParameters)});
    end
    

    figure(B2)
    children = get(B2, 'children');
    for ii=1:length(children)
        myTag = get(children(ii),'Tag');
        display(myTag);
        if ~strcmp(myTag,'Colorbar')
            delete(children(ii));
        end
    end
    
    %% Plot colorbar to same page
    myArg.defaultSlideNumber = 'current';
    
    %%
    
    if strcmp(myArg.externalParameters,'')
        %pptfigure(B2,'AutoGroupMin',myArg.defaultAutoGroupMin,'SlideNumber','current');
        pptfigure(B2,myArg);
    else
        pptfigure(B2,myArg);
        %pptfigure(B2,'AutoGroupMin',myArg.defaultAutoGroupMin,'SlideNumber','current',myArg.externalParameters{1:length(myArg.externalParameters)});
    end

    
    %pptfigure(B2,'BitmapResolution',200,'SlideNumber','current');

elseif hasColorbar == 0
    if strcmp(myArg.externalParameters,'')
        pptfigure(H2,myArg);
        %%pptfigure(H2,'AutoGroupMin',myArg.defaultAutoGroupMin,'SlideNumber',myArg.defaultSlideNumber);
    else
        pptfigure(H2,myArg);
        %pptfigure(H2,'AutoGroupMin',myArg.defaultAutoGroupMin,'SlideNumber',myArg.defaultSlideNumber,myArg.externalParameters{1:length(myArg.externalParameters)});
    end
end



end