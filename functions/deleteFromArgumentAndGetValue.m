function out = deleteFromArgumentAndGetValue(arguments,searchString)

value        = [];

try
    indexSearch = findStringInStructure(arguments,searchString);

catch myError
    indexSearch = -1;
    %Notfound => thats even better ;-)
end

if indexSearch ~= -1
    %%%Yes the user setted slidenumber to an other value => the value is
    %%%saved at slidenumber+1
    
    value = arguments{indexSearch+1}; %%Overwrite default value with user input
        
    %Now lets delete that value
    arguments = deleteFromStructure(arguments,[indexSearch,indexSearch+1]);
end

out.value       = value;
out.arguments   = arguments;
out.indexSearch = indexSearch;

end