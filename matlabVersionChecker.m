%% This function will help us do identify if the Matlab current version is higher or equal to the minversion we need for a certain feature.

function versionStatus = matlabVersionChecker(minVersion)

    versionStatus = 0; %% 0 = curVersion is lower than minVersion, 1 = Version Match, 1 = curVersion is higher than minVersion


    % Get the current matlab version
    curVersion = version('-release');
    
    
    curVersionStruct = splitMatlabVersion(curVersion);
    minVersionStruct = splitMatlabVersion(minVersion);
    
    % Check releaseVersion
        if curVersionStruct.releaseYear > minVersionStruct.releaseYear
            versionStatus = 1;
        end
        
        if curVersionStruct.releaseYear < minVersionStruct.releaseYear
            versionStatus = 0;
        end
        
        if curVersionStruct.releaseYear == minVersionStruct.releaseYear
            if curVersionStruct.releaseVersion >= minVersionStruct.releaseVersion
                versionStatus = 1;
            end
        end
    % 
    
    
    % Nested helper function
    function curToken = splitMatlabVersion(myVersion)
        
        % A version tag is build up like year + Version e.g 2015a
        
        % Define regexp string
        regExpSting = '(?<releaseYearString>\d*)(?<releaseVersion>.*)';
    
        [curToken,~] = regexp(myVersion, regExpSting,'names');
        
        releaseYear = 0;
    
        % Cast year to num
        try
           releaseYear = str2num(curToken.releaseYearString);
        catch
            warning(['An unexpected error occured when analyizing releaseYear: ',curToken.releaseYear])
        end
        
        curToken.releaseYear    = releaseYear;

    end
    

end


