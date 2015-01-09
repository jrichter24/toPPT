function myText = toLinebreakVersion(text)
 myText = '';

 isCell       = 0;
 isPureString = 0;
 %% Check if cellarray or string
 
 if ischar(text) %The user wants to add text to the slide
     isPureString = 1;
     myText = text;
 end
 
 %testtext = ['First Item', char(13), 'Second Item', char(13) , 'Third Item'];%Working
 
 if iscell(text) %The user wants to add text to the slide
     isCell = 1;
     
     for ii=1:length(text)
         if ii ~= length(text)
            myText = [myText,text{ii},char(13)];
         else
            myText = [myText,text{ii}];
         end
     end
 end
 
 %% If nor text nor cell return a -1
 
 if isCell == 0 && isPureString == 0
    myText = -1;
 end
 
 
end