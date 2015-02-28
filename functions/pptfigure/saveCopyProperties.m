% This will try to copy as much properties as possible from sourceObj to
% targetObj

function saveCopyProperties(sourceObj,targetObj,corrector)

    if strcmp(version('-release'),'2014b')
        
        % Get properties from sourceObj
        propStruct = get(sourceObj);
        
        % Get fieldnames
        propFields = fieldnames(propStruct);
        
        
        for ii=1:numel(propFields)
            try
                % Allpy Corrector First
                
                % First is cuurent propFields{ii} in corrector?
                rowIndex = find(ismember(corrector,propFields{ii}),true);
                
                if ~isempty(rowIndex)
                    exchangeProp = corrector{rowIndex,2};
                else
                    exchangeProp = 'null';
                end
                
                if strcmp(exchangeProp,'')
                    %Do nothing
                elseif ~strcmp(exchangeProp,'') && ~strcmp(exchangeProp,'null')
                    
                    set(targetObj,exchangeProp,propStruct.(propFields{ii}));
                    %display(['Added - ',' Field: ',exchangeProp,' Value: ',propStruct.(propFields{ii})]);
                    
                else
                    
                    set(targetObj,propFields{ii},propStruct.(propFields{ii}));
                    %display(['Added - ',' Field: ',propFields{ii},' Value: ',propStruct.(propFields{ii})]);
                    
                end
                
            catch
                %display(['Not Added - ',' Field: ',propFields{ii}]);
            end
        end

    else
        warning('This function is only available for Matlab release 2014b')
    end

end