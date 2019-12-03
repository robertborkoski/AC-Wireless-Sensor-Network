%%..................................................................................%%
%....________.....ECE 301 - Circuits and Electromechanical Computing.................%
%....|/.||.\|.....Air Conditioning Monitoring System (ACMS Central Node).............%
%.......||........Authored by Robert Borkoski and Carter Sutton......................%
%......_/\_.......In collaboration with Blake Thompson and Cameron Davis.............%
%%................Revised December 1 2019...........................................%%

%%..Table of Contents...............................................................%%
%.+-------------------------+------------+...........................................%
%.|.Function................|.Line.......|...........................................%
%.+-------------------------+------------+...........................................%
%.|.main()..................|.25.........|...........................................%
%.|.ReadOutlook.............|.66.........|...........................................%
%.|.Vargarin................|.186........|...........................................%
%.|.ProblemCodeDecomp.......|.226........|...........................................%
%.|.ExecuteSolutions........|.273........|...........................................%
%.|.IdentifyRecipients......|.304........|...........................................%
%.|.SendEmailFromArduino....|.322........|...........................................%
%.|.TranlateHall............|.345........|...........................................%
%.|.GetMessage..............|.366........|...........................................%
%%..................................................................................%%

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function main()

%% initialization

clear all, close all, clc
Proc('init');
time = 0; % timer - every 10 mins, conditions are checked
eval = 0; % not evaluating sensors until time trigger (eval = 1 will exit waiting loop)
on = 1; % system running
global SetUpEmailProtocolYet, global code_id
SetUpEmailProtocolYet = 0; % needs to set up SMTP first time through
while on == 1
	tic % start timer

	%% wait 10 minutes, then check email

	while eval == 0;
        Proc('wait')
		pause(1); % wait 1 s
		time = toc; % write timer
		if time > 1 % after 1 s (will change to 10 min upon final testing)
			eval = 1; % check email
		else
			eval = 0; % keep waiting
		end
    end
	%% check email
	mails = ReadOutlook; % check email, pull in any incoming problems % testing - until sensor segment done
	Proc('sc')
	unit_conditions = ProblemCodeDecomp(mails); % decompose problem into identifiers
    Proc('pull')
	ExecuteSolutions(unit_conditions); % take any necessary action based on what the problem is (emails, etc)
	clear unit_conditions, clear mails;
    on = 0;
end

end % end main

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function [Email]= ReadOutlook(varargin)

%% Function Inputs
vargs = varargin;
f = Varargin(vargs);
clearvars varargin vargs

%% Connects to Outlook
Proc('auth')
outlook = actxserver('Outlook.Application');
mapi = outlook.GetNamespace('mapi');
INBOX = mapi.GetDefaultFolder(6);

%% Retrieving UnRead or read emails / save or not save attachments
if isempty(f.Folder) && isempty(f.Subfolder)
    % reads Inbox only
    count = INBOX.Item.Count;
    Email = cell(count,2);
elseif ~isempty(f.Folder)
    % reads Inbox folder
    folder_numbers = INBOX.Folders.Count;
    % find folder / subfolder's outlookindex
    for i = 1:folder_numbers
        name = INBOX.Folders(1).Item(i).Name;
        if strcmp(name,f.Folder)
            n = i;
        end
    end    
    switch f.Subfolder
        % working for folder emails
        case ''
            % number of emails
            count = INBOX.Folders(1).Item(n).Items.Count;
            % cell for emailbody
            Email = cell(count,2);
        otherwise
            % Search for nth Inbox folder and count sub-folders
            folder_numbers = INBOX.Folders(1).Item(n).Folders(1).Count;
            % find Outlook Subfolder Index
            for i=1:folder_numbers
                name = INBOX.Folders(1).Item(n).Folders(1).Item(i).Name;
                if strcmp(name,f.Subfolder)
                    s= i;
                end
            end
            % number of emails
            count = INBOX.Folders(1).Item(n).Folders(1).Item(s).Items.Count;
            % cell for emailbody
            Email = cell(count,2);
    end
end

%% download & read emails
for i = 1:count
    if f.Read == 1 % only unreads emails
        % inbox
        if isempty(f.Folder) && isempty(f.Subfolder)
            UnRead = INBOX.Items.Item(count+1-i).UnRead;
        % folder
        elseif ~isempty(f.Folder) && isempty(f.Subfolder)
            UnRead = INBOX.Folders(1).Item(n).Items(1).Item(count+1-i).UnRead;
        % subfolder
        elseif ~isempty(f.Folder) && ~isempty(f.Subfolder)
            UnRead = INBOX.Folders(1).Item(n).Folders(1).Item(s).Item(1).Item(count+1-i).UnRead;
        end
        
        if UnRead
            % inbox
            if isempty(f.Folder) && isempty(f.Subfolder)
                if f.Mark == 1
                INBOX.Items.Item(count+1-i).UnRead=0;
                end
                email = INBOX.Items.Item(count+1-i);
                % folder
            elseif   ~isempty(f.Folder) && isempty(f.Subfolder)
                if Mark == 1
                INBOX.Folders(1).Item(n).Items(1).Item(count+1-i).UnRead=0;
                end
                email = INBOX.Folders(1).Item(n).Items(1).Item(count+1-i);
                % subfolder
            elseif ~isempty(f.Folder) && ~isempty(f.Subfolder)
                if f.Mark == 1
                INBOX.Folders(1).Item(n).Folders(1).Item(s).Item(1).Item(count+1-i).UnRead=0;
                end
                email = INBOX.Folders(1).Item(n).Folders(1).Item(s).Items.Item(count+1-i);
            end
        end
    else   % all emails
        % inbox
        if isempty(f.Folder) && isempty(f.Subfolder)
            email = INBOX.Items.Item(count+1-i);
            % folder
        elseif   ~isempty(f.Folder) && isempty(f.Subfolder)
            email = INBOX.Folders(1).Item(n).Items(1).Item(count+1-i);
            % subfolder
        elseif ~isempty(f.Folder) && ~isempty(f.Subfolder)
            email = INBOX.Folders(1).Item(n).Folders(1).Item(s).Items.Item(count+1-i);
        end
        UnRead = 1; %pseudo for next step
    end
    if UnRead
        % read and save body
        subject = email.get('Subject');
        body = email.get('Body');
        Email{i,1} = subject;
        Email{i,2} = body;
        if ~isempty(f.Savepath)
            attachments = email.get('Attachments');
            if attachments.Count >= 1
                fname = attachments.Item(1).Filename;
                full = [f.Savepath,'\',fname];
                attachments.Item(1).SaveAsFile(full)
            end
        end
    end
end
Email(all(cellfun('isempty', Email),2),:)=[];
end % end ReadOutlook

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function f = Varargin(vargs) % takes function inputs and rewrites as data structure
% varargin as structure
n = length(vargs);
if n>0
    names = vargs(1:2:n);
    values = vargs(2:2:n);
    for ix=1:(n/2)
        f.(names{ix}) = values{ix};
    end
    
    if ~isfield(f, 'Folder')
        f.Folder = '';
    end
    
    if ~isfield(f, 'Subfolder')
        f.Subfolder = '';
    end
    
    if ~isfield(f, 'Savepath')
        f.Savepath = '';
    end
    
    if ~isfield(f, 'Read')
        f.Read = '';
    end
    
    if ~isfield(f, 'Mark')
        f.Mark = '';
    end
else
    f.Folder = '';
    f.Subfolder = '';
    f.Savepath = '';
    f.Read = '';
    f.Mark = '';
end
end % end Vargarin

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function [unit_conditions] = ProblemCodeDecomp(mails)

%% initialization


%% example of cell array data structure to be implemented - room data modified as necessary as problems come in
% Orange_hall = zeros(100,5)
% Orange_Hall_rooms = {'107';'108';'113';'114';'115';'116';'117';'119';'121';'123';'124';'125';'126';'127';'128';'129';'130';'131';'136A';'136B';'136C';'136D';'137A';'137B';'137C';'137D';'142';'143';'144';'145';'150';'151';'152';'154';'156';'158';'159';'160';'161';'162';'163';'164';'165';'166';'171A';'171B';'171C';'171D';'172A';'172B';'172C';'172D'}
% Orange_Hall = mat2cell(Orange_Hall)
%%

%% from development - explains format of incoming data string from nodes and what variables they are assigned to
% Need:
%  - Hall ID
%  - Room ID
%  - Temp sensor status code
%  - Water sensor status code
%  - Humidity sensor code
%  - "Okay" status code - for future development
% format ex
% hall_room_temp_water_hum_okay = 01_121_0_0_1_0 ==> hall = 01, room = 121, temp = 0, water = 0, hum = 1, okay = 0
%%

%% Check email

%% Grab problem code from email and generate cell array unit_conditions
global code_id
okay_status = 1; % for later development - node confirms status with central node once a day (indicates that the node can communicate with the server)
code_id = mails{1,2}; % mails (output of ReadOutlook) is 2 row data structure - header in first row, problem code in second row
unit_conditions = strsplit(code_id,'_'); % turns 1 element problem code into a series of data strings to be assigned to different variables below (for action)
unit_conditions_headers = {'hall_id','room_id','temp_status','water_status','hum_status','okay_status'}; % input headers of cell array vector - now obsolete (but do not remove)
unit_conditions = vertcat(unit_conditions_headers,unit_conditions); % form mobile data vector

%% form identifiers / problem statuses from original input string - all in string format bc rooms need to be able to input char
hall_id = unit_conditions{2,1}; % Orange Hall, North Carrick, etc
room_id = unit_conditions{2,2}; % 121, 137A, etc
temp_status = unit_conditions{2,3}; % temp too high / low = 1, normal = 0
water_status = unit_conditions{2,4}; % water leak = 1, normal = 0
hum_status = unit_conditions{2,5}; % humidity too high / low = 1, normal = 0
okay_status = unit_conditions{2,6}; % for later development - node confirms status with central node once a day (indicates that the node can communicate with the server)

end % end ProblemCodeDecomp

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function ExecuteSolutions(unit_conditions)
global SetUpEmailProtocolYet
on = 1; % reduces time spent evaluating conditions when not necessary
counter = 2; % start at top of rooms in list - do not include header
hall_id = unit_conditions{2,1};
room_id = unit_conditions{2,2};
dim = size(unit_conditions); % find # of rooms in hall array
while on == 1
    if strcmp(unit_conditions{counter,6},"1") == 1 % okay_status checker - for future testing (see ProblemCodeDecomp)
        on = 0; % stop evaluating rooms - saves time if the only signal is the daily check in
    elseif strcmp(unit_conditions{counter,3},'1') == 1 % if temp is the problem
        Proc('ist')
        hall_id = unit_conditions{counter, 1}; % grab hall
        room_id = unit_conditions{counter, 2}; % grab room
        problem_id = 3; % log problem in room conditions database, under temperature condition
        recipients = IdentifyRecipients(hall_id, room_id, problem_id);
        SendEmailFromArduino(recipients,hall_id, room_id, problem_id); % 
    elseif strcmp(unit_conditions{counter,4},'1') == 1 
        Proc('isw')
        hall_id = unit_conditions{counter, 1}; % grab hall
        room_id = unit_conditions{counter, 2}; % grab room
        problem_id = 4; % log problem in room conditions database, under water condition
        recipients = IdentifyRecipients(hall_id, room_id, problem_id);
        SendEmailFromArduino(recipients,hall_id, room_id, problem_id);
    elseif strcmp(unit_conditions{counter,5},'1') == 1
        Proc('ish')
        hall_id = unit_conditions{counter, 1}; % grab hall
        room_id = unit_conditions{counter, 2}; % grab room
        problem_id = 5; % log problem in room conditions database, under humidity condition
        recipients = IdentifyRecipients(hall_id, room_id, problem_id);
        SendEmailFromArduino(recipients,hall_id, room_id, problem_id);
    end
on = 0;
end
counter = counter + 1; % next room in list
end % end ExecuteSolutions
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function [recipients] = IdentifyRecipients(hall_id, room_id, problem_id)


%% define available personnel to respond to maintenance concerns - includes hall directors, maintenance staff, maintenance supervisor
hall_personnel = {'Hall Name' , 'Orange' , 'Magnolia / Dogwood' , 'North Carrick' , 'South Carrick' , 'Reese' ; 'Hall Director' , 'Yasja Hemmings' , 'Brendan Miller' , 'Jordan Prewitt' , 'Kelly Gilton' , 'Claire Chernowsky' ; 'Email' , 'rborkosk@vols.utk.edu' , 'rborkosk@vols.utk.edu' , 'rborkosk@vols.utk.edu' , 'rborkosk@vols.utk.edu' , 'rborkosk@vols.utk.edu' ; 'facilities personnel' , 'Joe' , 'Jim' , 'Bob' , 'Billy' , 'Mac' ; 'Email' , 'robertborkoski@gmail.com' , 'robertborkoski@gmail.com' , 'robertborkoski@gmail.com' , 'robertborkoski@gmail.com' , 'robertborkoski@gmail.com'};
facilities_director = {'Name' , 'Mike West' , 'rborkosk@vols.utk.edu'};
%% based on hall_id, identify necessary personnel to be notified
hall_director_email = hall_personnel{3,str2num(hall_id) + 1}; % hall director
facilities_personnel_email = hall_personnel{5,str2num(hall_id) + 1}; % maintenance staff
facilities_director_email = facilities_director{1,3}; % facilities supervisor

recipients = {hall_director_email,facilities_personnel_email,facilities_director_email}; % wrap emails into one output variable data bus

end % end IdentifyRecipients

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function SendEmailFromArduino(recipients, hall_id, room_id, problem_id)
global SetUpEmailProtocolYet
if SetUpEmailProtocolYet == 0;
    setpref('Internet','SMTP_Server','smtp.gmail.com'); % define protocol to be used - SMTP server set up
    setpref('Internet','E_mail','ut.housing.ac@gmail.com'); % email address to send from
    setpref('Internet','SMTP_Username','ut.housing.ac@gmail.com'); % SMTP login info - username
    setpref('Internet','SMTP_Password','UTHousingSucks'); % SMTP login info - password
    props = java.lang.System.getProperties;
    props.setProperty('mail.smtp.auth','true'); % allow SMTP to send emails
    props.setProperty('mail.smtp.socketFactory.class', 'javax.net.ssl.SSLSocketFactory');
    props.setProperty('mail.smtp.socketFactory.port','465'); % server port to send email from - predefined 465
    SetUpEmailProtocolYet = 1;
end

counter = 1; % start with first email address

recipients = IdentifyRecipients(hall_id, room_id, problem_id)
MessageToSend = GetMessage(problem_id,hall_id,room_id)
for k = 1:3 % for all email addresses to send emails to
	sendmail(recipients{counter},['Maintenance issue detected - Room ' num2str(room_id)] ,MessageToSend{problem_id - 1 , counter + 1}); % send an email with the appropriate message (called from MessageToSend)
	counter = counter + 1; % go to next email address to send info to - NOTE: predefined order of sending emails to ensure proper message sent to each person (HD, then maint, then supervisor)
end

end % end SendEmailFromArduino

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function hall = TranslateHall(hall_id)
hall = 'undefined' % clear, introduce output variable hall, which will be part of email text
while strcmp(hall,'undefined') == 1 % reduces evaluating time - stops trying to figure out what hall it is when it is identified
	if strcmp(hall_id, '01') == 1 % for Orange Hall
		hall = 'Orange Hall';
	elseif strcmp(hall_id,'02') == 1 % for Magnolia / Dogwood Hall
		hall = 'Magnolia / Dogwood Hall';
	elseif strcmp(hall_id, '03') == 1 % for North Carrick Hall
		hall = 'North Carrick Hall';
	elseif strcmp(hall_id, '04') == 1 % for South Carrick Hall
		hall = 'South Carrick Hall';
	elseif strcmp(hall_id, '05') == 1 % for Reese Hall
		hall = 'Reese Hall'
	else
		hall = 'Error - invalid hall code' % email text will say invalid hall code - future development will include a troubleshooting program to determine why the problem code sent from the node was invalid
	end
end
end % end function TranslateHall

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function MessageToSend = GetMessage(problem_id,hall_id,room_id)


% differential text segments for emails
intro = 'This is a notification from the AC Monitoring system.  '
facilities_issue_identifier_temp = ['The temperature in ' TranslateHall(hall_id) ' room ' room_id ' is unusually high.  ']
facilities_issue_identifier_water = ['Sensors indicate an AC leak in ' TranslateHall(hall_id) 'room' room_id ', which needs to be addressed immediately.  ']
facilities_issue_identifier_hum = ['The humidity in ' TranslateHall(hall_id) ' room ' room_id ' is unusually high.  ']
facilities_action = ['Perform any required maintenance to the AC unit, and as soon as the request is completed, notify ut.housing.ac@outlook.com, including the following text: "AC maintenance completed ' TranslateHall(hall_id) ' ' 'room'  ' '  room_id '."']
HD_action = 'Please follow up with maintenance personnel if the issue is not resolved within 24 hours (or 72 hours if a weekend).'
supervisor_action = HD_action;

% messages template - overwritten with appropriate email text according to what type of problem it is - format will be redone in future development because it is clunky and obsolete but it works
messages = {'Problem type' , 'HD Message' , 'Facilities Personnel Message' , 'Facilities Director Message' ; 'Temperature' , 'HDM' , 'FPM' , 'FDM' ; 'Water' , 'HDM' , 'FPM' , 'FDM' ; 'Humidity' , 'HDM' , 'FPM' , 'FDM'}; % placeholder for email text



if problem_id == 3 % if temperature issue
	FacilitiesMessage = [intro facilities_issue_identifier_temp facilities_action] % email text to be sent to maintenance staff - temp
	HDMessage = [intro  facilities_issue_identifier_temp  HD_action] % email text to be sent to Hall Director - temp
	SupervisorMessage = HDMessage; % text will be replaced and improved in future development
	write = 2; % write email text to temperature row of messages array - passed into SendEmailToArduino
elseif problem_id == 4 % if humidity issue
	FacilitiesMessage = [intro facilities_issue_identifier_water facilities_action] % email text to be sent to maintenance staff - water
	HDMessage = [intro facilities_issue_identifier_water HD_action] % email text to be sent to Hall Director - water
	SupervisorMessage = HDMessage; % text will be replaced and improved in future development
	write = 3; % write email text to water row of messages array - passed into SendEmailToArduino
elseif problem_id == 5 % if humidity issue
	FacilitiesMessage = [intro facilities_issue_identifier_hum facilities_action] % email text to be sent to maintenance staff - humidity
	HDMessage = [intro facilities_issue_identifier_hum HD_action] % email text to be sent to Hall Director - humidity
	SupervisorMessage = HDMessage % text will be replaced and improved in future development
	write = 4; % write email text to humidity row of messages array - passed into SendEmailToArduino
end
if write == 2 || write == 3 || write == 4 % should never be denied if there is any text in HDMessage, FacilitiesMessage, or SupervisorMessage to write
	messages{write,2} = HDMessage % Hall director email text
	messages{write,3} = FacilitiesMessage % Maintenance staff email text
	messages{write,4} = SupervisorMessage % Facilities supervisor email text
end

MessageToSend = messages; % finished array - messages passed into SendEmailToArduino
clear FacilitiesMessage, clear HDMessage, clear SupervisorMessage, clear write % prevents text left over from last message from being sent incorrectly
end % function GetMessage

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

function Proc(step)
global code_id
if strcmp(step, 'init') == 1
    fprintf('Initializing.')
    pause(0.1)
    fprintf('.')
    pause(0.05)
    fprintf('.')
    pause(0.05)
    fprintf('.')
    pause(0.05)
    fprintf('.')
    pause(0.025)
    fprintf('.')
    pause(0.025)
    fprintf('.')
    fprintf('\n')
    fprintf('Completed Setup'),fprintf('\n')
elseif strcmp(step,'wait') == 1
    fprintf('Waiting for input (0)'),fprintf('\n')
    pause(1)
    fprintf('Problem received'),fprintf('\n')
elseif strcmp(step, 'auth') == 1
    fprintf('Authentication successful'),fprintf('\n')
    fprintf('Reading.')
    pause(0.05)
    fprintf('.')
    pause(0.025)
    fprintf('.'),fprintf('\n')
elseif strcmp(step,'pull') == 1
    fprintf('The following code(s) will be processed: '),fprintf(code_id),fprintf('\n')
elseif strcmp(step,'sc') == 1
    fprintf('Pulling 1 code(s) into Matlab'),fprintf('\n')
    fprintf('Action(s) needed to remediate: 3'),fprintf('\n')
elseif strcmp(step,'ist') == 1
    fprintf('issue type:  high temperature'),fprintf('\n')
    fprintf('Action(s) needed to remediate: 3'),fprintf('\n')
elseif strcmp(step,'isw') == 1
    fprintf('issue type:  water leak'),fprintf('\n')
    fprintf('Action(s) needed to remediate: 3'),fprintf('\n')
elseif strcmp(step,'ish') == 1
    fprintf('issue type: high humidity - WARNING: MOLD GROWTH LIKELY'),fprintf('\n')
end
end