function index = findStringInStructure(mystructure,mystring)

count = 0;
index = -1;

for ii=1:length(mystructure)
    try
    if strcmp(mystructure{ii},mystring)
        index(count+1) = ii;
        count = count+1;
    end
    catch
        % Seems not to be a string => No problem ;-)
    end
end

end