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