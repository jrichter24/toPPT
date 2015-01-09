function newstructure = deleteFromStructure(structure,arrayIndices)

newstructure = '';

jj = 1;
for ii =1:length(structure)
 
    if isempty(find(arrayIndices==ii, 1))
        newstructure{jj} = structure{ii};
        jj = jj+1;
    end
  
end
    
end