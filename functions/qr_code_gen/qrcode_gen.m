% qrcode_gen - 1.0
% ================
% Author :  Jens Richter
% eMail  :  jrichter@iph.rwth-aachen.de
% Date   :  February 2015
%
%
% Message parameter:
% ==================================
% The first parameter is the message that you want to encode.
% Syntax: 
%   qrcode_gen('My Message');
%
% Description:
%   Generates a QR code with your encoded message (with default parameters).
%
% Available parameters for qrcode_gen(input):
% ==================================
%
%
% Parameter 'Size':
% ---------------
% Syntax:
%   qrcode_gen(message, 'Size',[117,117]);
%   or equally:
%   qrcode_gen(message, 'Size',117);
%
% Description:
%   Generates a QR code of the defined Size. The size has to match the
%   17+4N rule (e.g. 17+4*25 = 117).
%
%
% Parameter 'Version':
% ---------------
% Syntax:
%   qrcode_gen(message, 'Version',10);
%
% Description:
%   Generates a QR code of the defined QR code version.
%
%
% Parameter 'ErrQuality':
% ---------------
% Syntax:
%   qrcode_gen(message, 'ErrQuality','L');
%
% Description:
%   ErrQuality accepts one of the following strings "L" or "M" or "H" or
%   "Q".
% Definition of error correction codes (taken from
% http://en.wikipedia.org/wiki/QR_code):
%   Level L (Low) 	7% of code words can be restored.
%   Level M (Medium) 	15% of code words can be restored.
%   Level Q (Quartile)[40] 	25% of code words can be restored.
%   Level H (High) 	30% of code words can be restored.
%
%
% Parameter 'CharacterSet':
% ---------------
% Syntax:
%   qrcode_gen(message, 'CharacterSet','UTF-8');
%
% Description:
%   Defines the character set that will be used. See qrcode_config for a
%   full list of available character sets.
%
%
% Parameter 'DownloadJars':
% ---------------
% Syntax:
%   qrcode_gen('DownloadJars', 1); - no message attribute is necessary
%
% Description:
%   This will download the necessary Jar files. So it is possible to use
%   qrcode_gen statically instead of importing the jars from the internet.
%
%
%
% Annotation:
% ---------------
% All parameters can be combined with each other e.g. 
%    qrcode_gen(message, 'ErrQuality','L','Version',10);

function qr = qrcode_gen(varargin)


%% First get settings from config
qrcodeProps = qrcode_config('qrcode');

%% Second check user inout
validateUserArguments(varargin);

%% Third load neccessary Jars

successloading = true;

loadJarsDynamically = false;
displayConole('Importing jars...');

    % First try to load jars statically from path
    for ii =1:numel(qrcodeProps.jar.files)
        pathLoad = [mfilename('fullpath'),qrcodeProps.jar.local.path];
        filePathLoad = [pathLoad,qrcodeProps.jar.files{ii},'-',qrcodeProps.jar.ver,'.jar'];
        
        if exist(filePathLoad, 'file')
            javaaddpath(filePathLoad);
            
        else
            loadJarsDynamically = true;
            break;
        end
        
    end
    displayConole(['Attempt to load jar files statically from: ',pathLoad]);
    
if loadJarsDynamically
    displayConole(['Loading jars statically failed...']);
    displayConole(['Attempt to load jar files dynamically from: ',qrcodeProps.jar.inet.path]);
    try
        % Second try to load from internet
        % Eg. http://repo1.maven.org/maven2/com/google/zxing/core/3.2.0/core-3.2.0.jar
        for ii=1:numel(qrcodeProps.jar.files)
            curPath = [qrcodeProps.jar.inet.path,qrcodeProps.jar.files{ii},'/',qrcodeProps.jar.ver,'/',qrcodeProps.jar.files{ii},'-',qrcodeProps.jar.ver,'.jar'];
            javaaddpath(curPath);
        end
         
    catch
        warning('Not able to load jars - stopping generation');
        successloading = false;
    end
end

%% If jars were loaded successfully go on!

if(successloading)
    
    displayConole('Loading jars...done');
    displayConole('Generating QR code...');
    
    try
        switch upper(qrcodeProps.setting.errorcQuality)
            case 'M'
                qr_quality = com.google.zxing.qrcode.decoder.ErrorCorrectionLevel.M;
            case 'L'
                qr_quality = com.google.zxing.qrcode.decoder.ErrorCorrectionLevel.L;
            case 'H'
                qr_quality = com.google.zxing.qrcode.decoder.ErrorCorrectionLevel.H;
            case 'Q'
                qr_quality = com.google.zxing.qrcode.decoder.ErrorCorrectionLevel.Q;
        end
        
        %% encoding qr
        qr_writer = com.google.zxing.qrcode.QRCodeWriter;
        qr_hints  = java.util.Hashtable;
        
        % use hint for encoding type
        qr_hints.put(com.google.zxing.EncodeHintType.ERROR_CORRECTION, qr_quality);
        
        % use hint for character set
        qr_hints.put(com.google.zxing.EncodeHintType.CHARACTER_SET, qrcodeProps.setting.characterset);
        
        M_java = qr_writer.encode(qrcodeProps.setting.message, com.google.zxing.BarcodeFormat.QR_CODE, qrcodeProps.setting.size(2) , qrcodeProps.setting.size(1) , qr_hints);
        qr = zeros(M_java.getHeight(), M_java.getWidth());
        
        for i=1:M_java.getHeight()
            for j=1:M_java.getWidth()
                qr(i,j) = M_java.get(j-1,i-1);
            end
        end
        
        clear qr_writer;
        clear M_java;
        
        qr = 1-logical(qr);
        displayConole('Generating QR code...done');
    catch
        warning('An Error occured during generating the QR-Code - This version of qrcode_gen is optimized for zxing 3.2.0');
        qr = 0;
    end
    
    
    
    
end

    function displayConole(consolemessage)
        if qrcodeProps.setting.suppressConsoleOut
            display(['',consolemessage]);
        end
    end

    function validateUserArguments(vararginRaw)
        
        if numel(vararginRaw)>0
            
            %; % Do nothing
            
        else %if is empty => set to ''
            vararginRaw    = cell(1);
            vararginRaw{1} = '';
        end
        
        % We are going to validate all arguments set by the user. If an
        % error happens we throw a warning and swith to the default value
        
        % Check for special commands
        checkSpecialCommandInput(vararginRaw);
        
        % Check version input
        checkVersionInput(vararginRaw);
        
        % Check error correction quality input
        checkErrorCorrectionQualityInput(vararginRaw);
        
        % Check size input
        checkQRCodeSizeInput(vararginRaw);

        % Check message input
        checkMessageInput(vararginRaw);
        
        % Check Character set input
        checkCharacterSetInput(vararginRaw);
        
    end

    function checkCharacterSetInput(vararginStruct)
        
        out = deleteFromArgumentAndGetValue(vararginStruct,'CharacterSet');
        if ~isempty(out.value)
            if ischar(out.value) && ismember(out.value,qrcodeProps.option.charactersets)
                
                %User setting is valid
                qrcodeProps.setting.characterset = out.value;
            else
                warning('The character set has to match the options defined in qrcode_config.');
            end
        end
    end

    function checkSpecialCommandInput(vararginStruct)
        
        out = deleteFromArgumentAndGetValue(vararginStruct,'DownloadJars');
        if ~isempty(out.value)
            % The user wants to download the jar files
            downloadJars();
        end
        
    end

    function downloadJars()
        
        % Second try to load from internet
        % Eg. http://repo1.maven.org/maven2/com/google/zxing/core/3.2.0/core-3.2.0.jar
        try
            displayConole('Creating folder structure for jars...');
            filePath = [mfilename('fullpath'),qrcodeProps.jar.local.path,'\'];
            mkdir(filePath);
            displayConole('Creating folder structure for jars...done');
            displayConole('Downloading jar files...');    
            for jj=1:numel(qrcodeProps.jar.files)
                
                curDownloadPath = [qrcodeProps.jar.inet.path,qrcodeProps.jar.files{jj},'/',qrcodeProps.jar.ver,'/',qrcodeProps.jar.files{jj},'-',qrcodeProps.jar.ver,'.jar'];
                saveToFileName = [filePath,qrcodeProps.jar.files{jj},'-',qrcodeProps.jar.ver,'.jar'];
                
                if strcmp(version('-release'),'2014b')
                    websave(saveToFileName,curDownloadPath);
                else
                    urlwrite(curDownloadPath,saveToFileName);
                end
            end
            displayConole('Downloading jar files...done');    
        catch
            warning('An error occured downloading the desired jar files.');
        end
        
    end

    function checkMessageInput(vararginStruct)
        
        messageChar = qrcodeProps.setting.message;
        
        if numel(vararginStruct) == 1 && strcmp(vararginStruct{1},'')
            warning('No message was specified for encoding - a test message will be used instead.');
            messageChar = qrcodeProps.setting.message;
        elseif numel(vararginStruct)>0
            % The first argument HAS to be the message
            message  = vararginStruct{1};
            if isnumeric(message)
                % Cast to char
                try
                    messageChar = num2str(message);
                catch
                    warning('Your message seems to be an numerical value but an error occured during casting to a string.');
                end
                
            elseif ischar(message)
                messageChar = message;
            else
                warning('The type of your message is unknown. Only numerical values and strings are allowed.');
            end
            
        end
        
        qrcodeProps.setting.message = messageChar;
    end

    function checkQRCodeSizeInput(vararginStruct)
        
        validInputSize = true;
        hasSizeInput   = false;
        
        out = deleteFromArgumentAndGetValue(vararginStruct,'Size');
        
        if qrcodeProps.setting.qrverUserAssigned
            
            temp = qrcodeProps.setting.initialSize+qrcodeProps.setting.quietZoneSize*qrcodeProps.setting.qrver;
            qrcodeProps.setting.size = [temp,temp];
            
        end
        
        if ~isempty(out.value) && ~qrcodeProps.setting.qrverUserAssigned
            hasSizeInput = true;
            
            if isnumeric(out.value)
                if numel(out.value) == 1
                    %Make a square sized qr codeout of scalar input
                    out.value = [out.value,out.value];
                end
                if numel(out.value) == 2
                    % check the size
                    if (out.value(1) == out.value(2)) % must be equal
                        tmp = out.value(1) - 17;
                        if not(mod(tmp,4) == 0)
                            warning('QR code size does not meet sizing requirements (17+N*4)');
                            validInputSize = false;
                        end
                    else
                        warning('QR code must be square defined as a square or matching sizing requirements (17+N*4)');
                        validInputSize = false;
                    end
                else
                    warning('Size needs to be a vector of two elements or a scalar element');
                    validInputSize = false;
                end
                
            else
                warning('QR code size needs to be specified by vector or scalar - e.g. "Size",[97,97] "');
                validInputSize = false;
            end
        elseif  ~isempty(out.value) && qrcodeProps.setting.qrverUserAssigned
            warning('A QR version was assigned. The version tag has priority over the size tag.');
        end
        
        if hasSizeInput && validInputSize %User setting is valid
           qrcodeProps.setting.size = out.value;
        end
        
    end

    function checkErrorCorrectionQualityInput(vararginStruct)
        
        out = deleteFromArgumentAndGetValue(vararginStruct,'ErrQuality');
        if ~isempty(out.value)
            if ischar(out.value) && ismember(out.value,qrcodeProps.option.errorcQualities)
                
                %User setting is valid
                qrcodeProps.setting.errorcQuality = out.value;
            else
                warning('The Error quality needs to be a String "L" or "M" or "H" or "Q"');
            end
        end
        
    end
    
    function checkVersionInput(vararginStruct)
        
        validVersionInput = false;
        qrcodeProps.setting.qrverUserAssigned = false;
        
        out = deleteFromArgumentAndGetValue(vararginStruct,'Version');
        if ~isempty(out.value)
            
            % Check validity of version argument - argument needs to be a
            % number beteen 1 and 40
            if isnumeric(out.value) && out.value>0 && out.value<=40
                validVersionInput = true;
            elseif isnumeric(out.value) && out.value<0 && out.value>40
                warning('Version has to be specified as number between 1 and 40 => switching to default version');
            elseif ischar(out.value)
                
                try
                    temp = str2double(out.value);
                    
                    if(temp>0 && temp<=40)
                        validVersionInput = true;
                    else
                        warning('Version has to be specified as number between 1 and 40 => switching to default version');
                    end
                catch
                    warning('Version has to be specified as number between 1 and 40 => switching to default version');
                end
                
                out.value = temp;
                
            else
                warning('Version has to be specified as number between 1 and 40 => switching to default version');
            end
            
            if validVersionInput
                % The user setted version is valid overwrite setting
                qrcodeProps.setting.qrver = out.value;
                qrcodeProps.setting.qrverUserAssigned = true;
            end
        end
        
    end

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
            try
                value = arguments{indexSearch+1}; %%Overwrite default value with user input
            catch
            end
            
            %Now lets delete that value
            arguments = deleteFromStructure(arguments,[indexSearch,indexSearch+1]);
        end
        
        out.value       = value;
        out.arguments   = arguments;
        out.indexSearch = indexSearch;
        
    end

    function index = findStringInStructure(mystructure,mystring)
        
        count = 0;
        index = -1;
        
        for jj=1:length(mystructure)
            try
                if strcmp(mystructure{jj},mystring)
                    index(count+1) = jj;
                    count = count+1;
                end
            catch
                % Seems not to be a string => No problem ;-)
            end
        end
        
    end

    function newstructure = deleteFromStructure(structure,arrayIndices)
        
        newstructure = '';
        
        jj = 1;
        for kk =1:length(structure)
            
            if isempty(find(arrayIndices==kk, 1))
                newstructure{jj} = structure{kk};
                jj = jj+1;
            end
            
        end
        
    end



end