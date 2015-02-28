% Help TEX2PPT:
%
% toPPT is able to translate TeX-code that it found in your text.
%
% Accepted TeX-code:
% ---------------
%
%     'alpha','upsilon','sim','beta','phi','leq','gamma','chi', ...
%     'infty','delta','psi','clubsuit','epsilon','omega','diamondsuit', ...
%     'zeta','Gamma','heartsuit','eta','Delta','spadesuit','theta', ...
%     'Theta','leftrightarrow','vartheta','Lambda','leftarrow','iota', ...
%     'Xi','uparrow','kappa','Pi','rightarrow','lambda','Sigma', ...
%     'downarrow','mu','Upsilon','circ','nu','Phi','pm','xi','Psi','geq', ...
%     'pi','Omega','propto','rho','forall','partial','sigma','exists', ...
%     'bullet','varsigma','ni','div','tau','cong','neq','equiv','approx', ...
%     'aleph','Im','Re','wp','otimes','oplus','oslash','cap','cup', ...
%     'supseteq','supset','subseteq','subset','int','in','o','rfloor', ...
%     'lceil','nabla','lfloor','cdot','ldots','perp','neg','prime', ...
%     'wedge','times','0','rceil','surd','mid','vee','varpi','copyright', ...
%     'langle','rangle','^','_','{','}','sqrt'
%
% Example:
% ---------------
%   toPPT(' x = \sqrt\x^2') - adds the formula to the current slide
%
%
% Interaction with basic and special css tags 
% => Use 'help addText' for more information about css-tags:
% ---------------
% 
%   toPPT('<i><b><s color:blue; font-size:40; font-family:Times New Roman>x = \sqrt\x^2<\s></b></i>');
%  
%   By default the css tags (<u>,<i>,<b> and <s>) have priority in combination with TeX tags
%   like color.
%   toPPT('TeX color will be overwritten by css tag: <s color:blue>\color{red}a^2+b^2 = c^2</s>');




function str = tex2ppt(range,txtString,onlyTranslate)


% txtString
[str, ppt] = tex2pptTranslate(txtString);


if(~onlyTranslate)
    set(range,'Text',str);
    for ch = 1:length(str)
        sel = range.Characters(ch);
        font = sel.Font;
        fontName = get(font,'Name');

        if ~isempty(ppt(ch).unicode)
            invoke(sel,'InsertSymbol',fontName, ...
                ppt(ch).unicode,1);
        end

        if ~isempty(ppt(ch).color)
            set(font.Color,'RGB', ...
                color2rgb(ppt(ch).color));
        end
        if ~isempty(ppt(ch).fontname)
            set(font,'Name',ppt(ch).fontname);
        end

        if ~isempty(ppt(ch).fontsize)
            scale = ppt(ch).fontsize / ...
                get(objects(obj),'fontsize');
            fontSize = get(font,'Size');
            set(font,'Size',fontSize*scale);
        end
        if ppt(ch).bold == 1
            set(font,'Bold','msoTrue');
        end
        if ppt(ch).italic == 1
            set(font,'Italic','msoTrue');
        end
        if ~isempty(ppt(ch).offset)
            totOff = length(ppt(ch).offset);
            vals = 1-0.6.^(0:totOff);
            vals = diff(vals);
            vals = vals.*ppt(ch).offset;
            bOff = sum(vals);
            set(font,'BaselineOffset',bOff);
        end
    end
end

function rgb = color2rgb(col)
% Converts Matlab 1x3 color vector to PowerPoint RGB number

col = round(col*255);
rgb = col(1)+256*col(2)+256^2*col(3);


function [str, ppt] = tex2pptTranslate(str)
% Converts TEX string into a formatted PowerPoint box

emptyPPT.unicode = [];
emptyPPT.bold = [];
emptyPPT.italic = [];
emptyPPT.color = [];
emptyPPT.fontname = [];
emptyPPT.fontsize = [];
emptyPPT.offset = [];
ppt = repmat(emptyPPT,1,length(str));

% Substitute backslash character
backSlash = regexp(str,'\\\\');
for indx = length(backSlash):-1:1
    str(backSlash(indx)+1) = [];
    ppt(backSlash(indx)+1) = [];    
    str(backSlash(indx)) = '-';
    ppt(backSlash(indx)).unicode = 92;
end

% Convert special characters to Unicode
TEX = {'alpha','upsilon','sim','beta','phi','leq','gamma','chi', ...
    'infty','delta','psi','clubsuit','epsilon','omega','diamondsuit', ...
    'zeta','Gamma','heartsuit','eta','Delta','spadesuit','theta', ...
    'Theta','leftrightarrow','vartheta','Lambda','leftarrow','iota', ...
    'Xi','uparrow','kappa','Pi','rightarrow','lambda','Sigma', ...
    'downarrow','mu','Upsilon','circ','nu','Phi','pm','xi','Psi','geq', ...
    'pi','Omega','propto','rho','forall','partial','sigma','exists', ...
    'bullet','varsigma','ni','div','tau','cong','neq','equiv','approx', ...
    'aleph','Im','Re','wp','otimes','oplus','oslash','cap','cup', ...
    'supseteq','supset','subseteq','subset','int','in','o','rfloor', ...
    'lceil','nabla','lfloor','cdot','ldots','perp','neg','prime', ...
    'wedge','times','0','rceil','surd','mid','vee','varpi','copyright', ...
    'langle','rangle','^','_','{','}','sqrt'};
code = [945 965 126 946 966 8804 947 967 8734 948 968 9827 603 969 9830 ...
    950 915 9829 951 916 9824 952 920 8596 977 923 8592 953 926 8593 ...
    954 928 8594 955 931 8595 181 978 186 957 934 177 958 936 8805 960 ...
    937 8733 961 8704 8706 963 8707 8226 962 8717 247 964 8773 8800 ...
    8801 8776 8501 8465 8476 8472 8855 8853 8709 8745 8746 8839 8835 ...
    8838 8834 8747 8712 959 8971 8968 8711 8970 183 8230 8869 172 180 ...
    8743 215 8709 8969 8730 124 8744 982 169 10216 10217 94 95 123 125 8730];

for c = 1:length(TEX)
    [st en] = regexp(str,['\\' TEX{c}]);
    for indx = length(st):-1:1
        str(st(indx)+1:en(indx)) = [];
        ppt(st(indx)+1:en(indx)) = [];
        str(st(indx)) = '-';
        ppt(st(indx)).unicode = code(c);
    end
end

% Identify colors
[st en] = regexp(str,'\\color\{[^[\{\}]]*\}');
for indx = length(st):-1:1
    ppt(st(indx)).color = str(st(indx)+7:en(indx)-1);
    str(st(indx)+1:en(indx)) = [];
    ppt(st(indx)+1:en(indx)) = [];
end

% Identify RGB colors
[st en] = regexp(str,'\\color\[rgb\]\{[^[\{\}\[\]]]*\}');
for indx = length(st):-1:1
    ppt(st(indx)).color = str(st(indx)+12:en(indx)-1);
    str(st(indx)+1:en(indx)) = [];
    ppt(st(indx)+1:en(indx)) = [];
end

% Fix colors
for c = 1:length(ppt)
    if isempty(ppt(c).color)
        continue
    end
    val = str2num(ppt(c).color); %#ok<ST2NM>
    if ~isempty(val)
        ppt(c).color = val;
    else
        switch ppt(c).color
            case 'black'
                col = [0 0 0];
            case 'red'
                col = [1 0 0];
            case 'green'
                col = [0 1 0];
            case 'blue'
                col = [0 0 1];
            case 'yellow'
                col = [1 1 0];
            case 'magenta'
                col = [1 0 1];
            case 'cyan'
                col = [0 1 1];
            case 'white'
                col = [1 1 1];
        end
        ppt(c).color = col;
    end
end

% Identify font names
[st en] = regexp(str,'\\fontname\{[^[\{\}\[\]]]*\}');
for indx = length(st):-1:1
    ppt(st(indx)).fontname = str(st(indx)+10:en(indx)-1);
    str(st(indx)+1:en(indx)) = [];
    ppt(st(indx)+1:en(indx)) = [];
end

% Identify font sizes
[st en] = regexp(str,'\\fontsize\{[^[\{\}\[\]]]*\}');
for indx = length(st):-1:1
    ppt(st(indx)).fontsize = str2double(str(st(indx)+10:en(indx)-1));
    str(st(indx)+1:en(indx)) = [];
    ppt(st(indx)+1:en(indx)) = [];
end

% Identify bold formatting
st = strfind(str,'\bf');
for indx = length(st):-1:1
    str(st(indx)+1:st(indx)+2) = [];
    ppt(st(indx)+1:st(indx)+2) = [];
    ppt(st(indx)).bold = 1;
end

% Identify italic formatting
st = strfind(str,'\it');
for indx = length(st):-1:1
    str(st(indx)+1:st(indx)+2) = [];
    ppt(st(indx)+1:st(indx)+2) = [];
    ppt(st(indx)).italic = 1;
end
st = strfind(str,'\sl');
for indx = length(st):-1:1
    str(st(indx)+1:st(indx)+2) = [];
    ppt(st(indx)+1:st(indx)+2) = [];
    ppt(st(indx)).italic = 1;
end

% Identify normal formatting
st = strfind(str,'\rm');
for indx = length(st):-1:1
    str(st(indx)+1:st(indx)+2) = [];
    ppt(st(indx)+1:st(indx)+2) = [];
    ppt(st(indx)).bold = 0;
    ppt(st(indx)).italic = 0;
end

% Encode subscript and superscript symbols
f = strfind(str,'^');
for c = length(f):-1:1
    if strcmp(str(f(c)+1),'{')
        str(f(c):f(c)+1) = '{\';
    else
        str = [str(1:f(c)-1) '{\' str(f(c)+1) '}' str(f(c)+2:end)];
        ppt = [ppt(1:f(c)-1) emptyPPT ppt(f(c):f(c)+1) ...
            emptyPPT ppt(f(c)+2:end)];
    end
    ppt(f(c)+1).offset = 1;
end
f = strfind(str,'_');
for c = length(f):-1:1
    if strcmp(str(f(c)+1),'{')
        str(f(c):f(c)+1) = '{\';
    else
        str = [str(1:f(c)-1) '{\' str(f(c)+1) '}' str(f(c)+2:end)];
        ppt = [ppt(1:f(c)-1) emptyPPT ppt(f(c):f(c)+1) ...
            emptyPPT ppt(f(c)+2:end)];
    end
    ppt(f(c)+1).offset = -1;
end

% Apply TEX formatting
[str, ppt, ~] = apply_tex_format(str, ppt, emptyPPT);


function [str, ppt, emptyPPT] = apply_tex_format(str, ppt, emptyPPT)

f = strfind(str,'{');
while ~isempty(f)
    f = f(1);
    % Find matching brace
    isBrace = zeros(size(str));
    isBrace(strfind(str,'{')) = 1;
    isBrace(strfind(str,'}')) = -1;
    level = cumsum(isBrace);
    level(1:f-1) = inf;
    match = find(level==0,1);
    [strN, pptN, emptyPPT] = apply_tex_format(str(f+1:match-1), ...
        ppt(f+1:match-1),emptyPPT);
    str = [str(1:f-1) strN str(match+1:end)];
    ppt = [ppt(1:f-1) pptN ppt(match+1:end)];
    f = strfind(str,'{');
end

% No more braces
currForm = emptyPPT;
for c = 1:length(str)
    if strcmp(str(c),'\')
        if ~isempty(ppt(c).color)
            currForm.color = ppt(c).color;
        end
        if ~isempty(ppt(c).fontname)
            currForm.fontname = ppt(c).fontname;
        end
        if ~isempty(ppt(c).fontsize)
            currForm.fontsize = ppt(c).fontsize;
        end
        if ~isempty(ppt(c).bold)
            currForm.bold = ppt(c).bold;
        end
        if ~isempty(ppt(c).italic)
            currForm.italic = ppt(c).italic;
        end
        if ~isempty(ppt(c).offset)
            currForm.offset = ppt(c).offset;
        end
    else
        if isempty(ppt(c).color)
            ppt(c).color = currForm.color;
        end
        if isempty(ppt(c).fontname)
            ppt(c).fontname = currForm.fontname;
        end
        if isempty(ppt(c).fontsize)
            ppt(c).fontsize = currForm.fontsize;
        end
        if isempty(ppt(c).bold)
            ppt(c).bold = currForm.bold;
        end
        if isempty(ppt(c).italic)
            ppt(c).italic = currForm.italic;
        end
        ppt(c).offset = [currForm.offset ppt(c).offset];
    end
end

% Remove backslashes
f = strfind(str,'\');
str(f) = [];
ppt(f) = [];