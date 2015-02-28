function figure1 = example_helper(figure_index)

switch figure_index
    case 1
        n = 10;
        r = (0:n)'/n;
        theta = pi*(-n:n)/n;
        X = r*cos(theta);
        Y = r*sin(theta);
        C = r*cos(2*theta);

        figure1 = figure;
        axes1 = axes('Parent',figure1,'PlotBoxAspectRatio',[1 1 1],...
            'DataAspectRatio',[1 1 1]);
        box(axes1,'on');
        hold(axes1,'all');
        pcolor(X,Y,C,'Parent',axes1)
        axis equal tight 
        cb = colorbar('peer',axes1);
        title('MyTestFor toPPT');
    case 2
        
        figure1 = figure;
        hold all
        x=0:0.1:2*pi;
        b=zeros(5);
        c={'one','two','three','four','five'};
        xlabel('x')
        ylabel('y')
        for a=1:3
            b(a)=plot(x,sin(a*x),'LineWidth',2,'LineStyle','-');
            legend(b(1:a),c(1:a),'FontSize',14,'FontName','High Tower Text')
            %pause(1)
        end
        title('My Figure for toPPT');
        
    case 3
        n = 10;
        r = (0:n)'/n;
        theta = pi*(-n:n)/n;
        X = r*cos(theta);
        Y = r*sin(theta);
        C = r*cos(2*theta);

        figure1 = figure;
        axes1 = axes('Parent',figure1,'PlotBoxAspectRatio',[1 1 1],...
            'DataAspectRatio',[1 1 1]);
        box(axes1,'on');
        hold(axes1,'all');
        pcolor(X,Y,C,'Parent',axes1)
        axis equal tight 
        title('My Figure for toPPT');
end


end