clear;
close all;

FontSize = 1:2:20;

myText = {'1','12','123','1234','12345','123456','12345678','1234567812 34567812345678'};
%myText = {'12345678','1234567812 34567812345678'};
myText = {'1','12','123','1234','12345','123456','12345678','123456789',...
    '1234567891','12345678912','123456789123','1234567891234','1234567891235','1234567891236'...
    ,'12345678912367','123456789123678','1234567891236789'};

numberOfChar = 10:10:50;

for ii = 1:numel(numberOfChar)
    tempString = '';
    for jj = 1:numberOfChar(ii)
        tempString = [tempString,'1'] ;
    end
    myText{ii} = tempString;
end


for kk=1:numel(myText)
    %% Generate figures
    for ii=1:numel(FontSize)
        
        figure1 = figure;
        hold all
        x=0:0.1:2*pi;
        b=zeros(5);
        c={'one',myText{kk},'three','four','five'};
        xlabel('x')
        ylabel('y')
        for a=1:3
            b(a)=plot(x,sin(a*x),'Marker','+','LineWidth',2,'LineStyle',':');
            legend(b(1:a),c(1:a),'FontSize',FontSize(ii),'FontName','High Tower Text')
            %pause(1)
        end
        title('My Figure for toPPT');
        
        
        
        fig1childs = get(figure1,'children');
        legendchilds = get(fig1childs(1),'children');
        %legendCertainChild = get(legendchilds(6));
        
        legendCertainPostition  = get(fig1childs(1),'Position');
        legendCertainChildPostition  = get(legendchilds(2),'XData');
        
        
        AxesPos3(ii) = legendCertainPostition(3);
        AxesPos4(ii) = legendCertainPostition(4);
        
        Pos1(ii) = legendCertainChildPostition(1);
        Pos2(ii) = legendCertainChildPostition(2);
        %display(get(legendchilds(6),'String'));
        display(legendCertainChildPostition(1));
        

        close all;
        
    end
    
    curTester1 = AxesPos3.*Pos1;
    testerMean1(kk) = mean(curTester1);
    testerVar1(kk) = var(curTester1);
    
    curTester2 = AxesPos3.*Pos2;
    testerMean2(kk) = mean(curTester2);
    testerVar2(kk) = var(curTester2);
    
    % figureRes = figure;
    % plot(FontSize,Pos1)
    
    [xData, yData] = prepareCurveData( FontSize, Pos1 );
    
    % Set up fittype and options.
    ft = fittype( 'poly2' );
    
    % Fit model to data.
    [fitresult, gof] = fit( xData, yData, ft );
    
    resP1(kk) = fitresult.p1;
    resP2(kk) = fitresult.p2;
    resP3(kk) = fitresult.p3;
    
    % Plot fit with data.
    figure( 'Name', 'untitled fit 1' );
    h = plot( fitresult, xData, yData );
    legend( h, 'Pos1 vs. FontSize', 'untitled fit 1', 'Location', 'NorthEast' );
    % Label axes
    xlabel( 'FontSize' );
    ylabel( 'Pos1' );
    grid on
    
 
    

    [xData, yData] = prepareCurveData( FontSize, AxesPos3 );
    
    % Set up fittype and options.
    ft = fittype( 'poly2' );
    
    % Fit model to data.
    [fitresult, gof] = fit( xData, yData, ft );
    
    resAP1(kk) = fitresult.p1;
    resAP2(kk) = fitresult.p2;
    resAP3(kk) = fitresult.p3;
    
    figure( 'Name', 'untitled fit 1' );
    h = plot( fitresult, xData, yData );
    legend( h, 'AxesPos3 vs. FontSize', 'untitled fit 1', 'Location', 'NorthEast' );
    % Label axes
    xlabel( 'FontSize' );
    ylabel( 'Pos1' );
    grid on
    
    
    close all;
    
end
% 
% figureRes = figure;
% plot(1:numel(myText),resP1) % P1 is constant
% hold
% plot(1:numel(myText),resP2,'r')
% plot(1:numel(myText),resP3,'black')
% 
% figureResA = figure;
% plot(1:numel(myText),resAP1,'b--') % P1 is constant
% hold
% plot(1:numel(myText),resAP2,'r--')
% plot(1:numel(myText),resAP3,'black--')
% 
% figureResDivided = figure;
% plot(1:numel(myText),resAP3.*resP3,'black--')

figureTester1 = figure;
plot(1:numel(myText),testerMean1)

figureTesterVar1 = figure;
plot(1:numel(myText),testerVar1)

figureTester2 = figure;
plot(1:numel(myText),testerMean2)

figureTesterVar2 = figure;
plot(1:numel(myText),testerVar2)