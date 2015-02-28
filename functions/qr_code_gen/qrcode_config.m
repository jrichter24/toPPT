% qrcode_config - 1.0
% ================
% Author :  Jens Richter
% eMail  :  jrichter@iph.rwth-aachen.de
% Date   :  February 2015

function desiredProperty = qrcode_config(identifierProperty)

    %% Jar and import config - representing the version, the dowload path
    % and the jar files
    qrcode.jar.inet.path  = 'http://repo1.maven.org/maven2/com/google/zxing/';
    qrcode.jar.local.path = '\jarfiles\';
    qrcode.jar.ver  = '3.2.0';
    qrcode.jar.files = {'core','javase'};
    qrcode.jar.importpackages = {'com.google.zxing.qrcode.QRCodeWriter','com.google.zxing.BarcodeFormat','com.google.zxing.EncodeHintType','com.google.zxing.qrcode.decoder.ErrorCorrectionLevel'};
    
    %% QR-Code options
    qrcode.option.errorcQualities = {'L' 'M' 'Q' 'H'};

        %  taken from http://github.com/zxing/zxing/blob/master/core/src/main/java/com/google/zxing/common/CharacterSetECI.java
    qrcode.option.charactersets ={'ISO-8859-1','ISO-8859-2','ISO-8859-3','ISO-8859-4','ISO-8859-5','ISO-8859-6','ISO-8859-7','ISO-8859-8','ISO-8859-9',...
                                  'ISO-8859-10','ISO-8859-11','ISO-8859-13','ISO-8859-14','ISO-8859-15','ISO-8859-16','Shift_JIS','windows-1250',...
                                  'windows-1251','windows-1252','windows-1256','UTF-16BE','UTF-8','US-ASCII','GB2312','EUC_CN', 'GBK','EUC-KR'};
    
    %% Default settings
    qrcode.setting.message        = 'This is a test message that is only thrown if no message was specified or the message you specified was not encodable. Check your console for warnings.';
    qrcode.setting.characterset   = 'UTF-8';
    qrcode.setting.qrver          = 1;    % Version of QR-Code (between 1 and 40)
    qrcode.setting.initialSize    = 17;   % minimal size of the QR code
    qrcode.setting.quietZoneSize  = 4;
    qrcode.setting.errorcQuality  = qrcode.option.errorcQualities{2} ; % error correction quality
    qrcode.setting.size           = [qrcode.setting.initialSize+qrcode.setting.quietZoneSize*qrcode.setting.qrver,qrcode.setting.initialSize+qrcode.setting.quietZoneSize*qrcode.setting.qrver];
    qrcode.setting.suppressConsoleOut = false; % If true only warnings are thrown via the conole out

    
    %% Now lets return the desired property
    switch identifierProperty
        case 'qrcode'
            desiredProperty  = qrcode;
    end
end