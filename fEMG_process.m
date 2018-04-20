%% Psychophys (EMG) Signal Standardization Script
% Evan Carr & Galit Hofree - Winkielman Lab
% UC San Diego (2013)

%---------------------------------------------%

clear all;

%% COLLECT DATA INFO

% Constants ...
dataSheet = 'Interval Stats';
triggerSheet = 'EMG Stats';
triggerRow = 42;

% Set up default values ...
defaultSetupFile = 'setupInfo';
defaultInterval = 500;
defaultDataType = 'mean';
defaultBaseType = 'trial';
defaultShortBase = 0;

% Creates several prompts to collect user input:
% 1st, see if user wants to run with a setup cell array of input ...
quickSetup = questdlg('Do you want to use an already setup .mat file with all the input information? Type "yes" or "no." ');
if (strcmpi(quickSetup,'yes'))
    
    inputMatName = inputdlg({'name of mat file:'},'Setup File',1,{defaultSetupFile});
    try
        load(inputMatName{1,1});
    catch exception
        errordlg('Setup file not found!');
        break;
    end
    
elseif (strcmpi(quickSetup,'cancel'))
    break;
else

    % 2nd, collect subject info ...
    subjectPrompt = {'Number of first subject: ','Number of last subject: '};
    subjTitle = 'Subject info';
    subjInfo = inputdlg(subjectPrompt,subjTitle);
    
    num1stSubj = str2double(subjInfo{1,1});
    numlastSubj = str2double(subjInfo{2,1});

    % 3rd, collect muscle info ...
    musclePrompt = {'Integral, Mean, or RMS? '};
    muscleTitle = 'Muscle info';
    muscleInfo = inputdlg(musclePrompt,muscleTitle,1,{defaultDataType});
    
    dataType = muscleInfo{1,1};

    % 4th, collect trial info ...
    trialPrompt = {'Length of trial in milliseconds: ','Length of grouping interval in milliseconds, if already used in MindWare: '};
    trialTitle = 'Trial info';
    trialInfo = inputdlg(trialPrompt,trialTitle,1,{'',num2str(defaultInterval)});
    
    trialLength = str2double(trialInfo{1,1});
    intervalLength = str2double(trialInfo{2,1});    
    num_totalvals_per_trial = trialLength/intervalLength;

    % 5th, collect baseline info ...
    basePrompt = {'Full length of baseline saved in MindWare (in milliseconds): ',...
        'If you''d like to use a shorter baseline than that saved, enter how much shorter in milliseconds (e.g., if you want to only use 500ms out of a 2000ms baseline, type ''1500''): ',...
        'Would you like to use a mean per trial, median of entire experiment baseline, sliding window, or minimum window baseline? (type ''trial'', ''median'', ''window'', or ''min''): '};
    baseTitle = 'Baseline info';
    baseInfo = inputdlg(basePrompt,baseTitle,1,{'',num2str(defaultShortBase),defaultBaseType});

    if ((strcmpi(baseInfo{3,1},'window') == 1) || strcmpi(baseInfo{3,1},'min') == 1)
        windowSize = inputdlg({'How many datapoints do you want to include before/after the datapoint in the sliding window? '},'Window size',1,{'20'});
        windowLength = str2double(windowSize{1,1});
    end
    
    baseFull = str2double(baseInfo{1,1});
    baseShort = str2double(baseInfo{2,1});
    baseType = baseInfo{3,1};
    
    num_baseline_per_trial = (baseFull-baseShort)/intervalLength;
    num_padding_per_trial = baseShort/intervalLength;
    
    % 6th, complete settings for outliers and data cleaning ...
    outlierDlg = questdlg('Do you want to clean outliers (Y / N)?', ...
                         'Data cleaning', ...
                         'Yes', 'No', 'Yes');
    
    if (strcmpi(outlierDlg, 'yes') == 1)
        outlierPrompt = {'Type the number of standard deviations from each subject''s mean you want set for outlier detection (e.g., 3) -- trial values greater or less than this threshold will be removed from subsequent analyses: '};
        outlierStdDevDlg = inputdlg(outlierPrompt, 'Outlier exclusion threshold', 1, {'3'});
        outlierSD = str2double(outlierStdDevDlg{1,1});
        outlierStr = ['Mean +/- ' num2str(outlierSD) 'SDs'];
    else outlierSD = 0;
        outlierStr = 'None';
    end
    
    zscoreDlg = questdlg('Do you want to zscore the output, or leave it raw?','Output','zscore','raw','zscore');
    
end % End of the user input GUI dialogs!

% Check if all the info is okay; if so, save to log / .mat file ...

if ((strcmpi(baseInfo{3,1},'window') == 1) || (strcmpi(baseInfo{3,1},'min') == 1))
        windowStr = ['Window size:' windowSize{1,1}];
else
    windowStr = '';
end

% Confirmation dialog for user ...
hCheck = questdlg({'Would you like to continue with this setup? ';...
'';...
['First subject: ' num2str(num1stSubj)]; ['Last subject: ' num2str(numlastSubj)];...
['Data type: ' dataType];...
['Trial length: ' num2str(trialLength)]; ['Interval length: ' num2str(intervalLength)];...
['Outlier exclusion threshold: ' outlierStr];...
['Original baseline: ' num2str(baseFull)]; ['Selected baseline: ' num2str((baseFull-baseShort))];...
['Baseline type: ' num2str(baseType)];
 windowStr;
 ['Outlier exclusion criteria: ', outlierStr];...
 ['Data output: ', zscoreDlg];
'';...
'If so, a .mat file will save these preferences.'},'Setup Details','Continue','Cancel','Continue');

% Save setup information into a mat file for reuse:
save('setupInfo');

% Use setup data as input into a log file for this session:
logfile = fopen(['EMG Log File ' datestr(now, 'D-dd-mm-yyyy-T-HH-MM') '.txt'],'w');
fprintf(logfile,'### INPUT DETAILS ###\r');

fprintf(logfile,'First subject: %d,',num1stSubj);
fprintf(logfile,'Last subject: %d\n',numlastSubj);
fprintf(logfile,'Data type: %s\n',dataType);
fprintf(logfile,'Trial length: %d\n',trialLength);
fprintf(logfile,'Interval length: %d\n',intervalLength);
fprintf(logfile,'Outlier exclusion threshold: %s\n',outlierStr);
fprintf(logfile,'Original baseline: %d\n',baseFull);
fprintf(logfile,'Selected baseline: %d\n',(baseFull-baseShort));
fprintf(logfile,'Baseline type: %s\n',baseType);
fprintf(logfile,'%s\n##########\n',windowStr);
fprintf(logfile,'Outlier exclusion type: %s\n',outlierStr);
fprintf(logfile,'Data output: %s\n',zscoreDlg);

%% START AUTOMATED ANALYSIS
% Open subject files 1-by-1 ...

curDir = dir('*.xlsx'); % grab the names of the xlsx files in the directory
numMuscles = [];
    
% Read in data from MindWare Excel file ...
for b = num1stSubj:numlastSubj
    
    % Find whether subject file exists, but allowing for various forms
    % of filenames that would include 'subject' and the number:
    i = 0;
    filename = '';
    while (i >= 0) && (i < size(curDir,1))
        i = i + 1;
        if(regexp(curDir(i).name,sprintf('[Ss]ubject\\s*(0)*%d(\\D+.*)*\\.xlsx',b))) % (Uses regular expressions to allow various filenames)
            filename = curDir(i).name;
            i = -1;
        end
    end
    
    if(isempty(filename)) % (Assumes some subjects might be missing)
        
        fprintf(logfile,'Missing Subject %d\n',b);
        
    else % Found the file!
        
        % Read data ...
        [numeric_data, txt_data] = xlsread(filename, dataSheet);
        [numeric_stats, txt_stats] = xlsread(filename, triggerSheet);
        
        % Create structure for all phases and DVs:
        if(isempty(numMuscles))
            muscleNames = unique(txt_data(2:end,4));
            numMuscles = length(muscleNames);
            
            % Log muscle info:
            fprintf(logfile,'Found %d muscles: ',numMuscles);
            muscVarNames = cell(4,1);
            for i = 1:(numMuscles-1)
                muscleName = char(muscleNames(i,1));
                muscVarName = genvarname(muscleName);
                fprintf(logfile,'%s,',muscleName);
                eval([muscVarName '_Data = [];']);
                muscVarNames{i,1} = muscVarName;
            end
            muscleName = char(muscleNames(numMuscles,1));
            muscVarName = genvarname(muscleName);
            fprintf(logfile,'%s\n',muscleName);
            eval([genvarname(muscleName) '_Data = [];']);
            muscVarNames{numMuscles,1} = muscVarName;
        end
        
        fprintf(logfile,'Found Subject %d ...\n',b);
        
        % Remove extra baseline rows (if baseline used is shorter than that in MindWare)
        % Also, check to see if there are any repetions of the last trial and remove them
        % (MindWare bug? Found several subjects with more than one last trial ...)
        totalTime = num_totalvals_per_trial+num_padding_per_trial+num_baseline_per_trial;
        lastTrialRows = find(numeric_data(:,1) == max(numeric_data(:,1)));
        if (lastTrialRows ~= (totalTime)*numMuscles)
            rowsToRemove = lastTrialRows((totalTime*numMuscles+1):end);
            numeric_data = removerows(numeric_data,rowsToRemove);
            txt_data(rowsToRemove,:) = [];
        end
        padRows = [];
        for i = 1: num_padding_per_trial
            padRows = [padRows, i:totalTime:size(numeric_data,1)];
        end
        numeric_data = removerows(numeric_data,padRows);
        txt_data((padRows+1),:) = [];

        % Create dataset ...
        format_data = dataset();

        format_data.Event_Number = numeric_data(:,1);
        format_data.Start_Time = numeric_data(:,2);
        format_data.End_Time = numeric_data(:,3);
        format_data.Channel_Name = txt_data(2:end,4);
        
        dataColumn = find(strcmpi(txt_data(1,:),dataType));
        format_data.EMG_data = numeric_data(:,dataColumn);

        % Extract trigger numbers: 
        triggerValues = txt_stats(triggerRow, 2:end);
        
        % Get number out of trigger text:
        [junk_info1, info1] = strtok(triggerValues, '#');
        [junk_info2, info2] = strtok(info1, ' ');
        triggerNumbers = zeros(size(info2,2),1);
        for i = 1: size(info2,2)
            triggerNumbers(i,1) = str2double(info2{i});
        end
        
        for i = 1:numMuscles
            muscleName = char(muscleNames(i,1));

            % Extract muscle info:
            muscleRows = find(strcmpi(format_data.Channel_Name,muscleName));

            % Create muscle data matrix without channel names for quicker
            % manipulations:
            muscleData = format_data.EMG_data(muscleRows);

            % Remove outliers:     
            % (reason for not using zscore function - does not work with
            % nan values)
            rawData = muscleData;
            if outlierSD > 0
                rawMean = nanmean(rawData);
                rawStd = nanstd(rawData);
                stdData = (rawData - rawMean)/rawStd;
                outliers = find(abs(stdData) > outlierSD);
                % Set outliers to be NaN (null):
                rawData(outliers) = NaN;
            end
            
            % Standardize (zscore):
            % (reason for not using zscore function - does not work with
            % nan values)
            if (strcmpi('zscore',zscoreDlg) == 1)
                cleanMean = nanmean(rawData);
                cleanStd = nanstd(rawData);
                outputData = (rawData - cleanMean)/cleanStd;
            else
                outputData = rawData;
            end
            

            % Arrange data in rows per trial ...
            orgData = reshape(outputData,num_totalvals_per_trial+num_baseline_per_trial,[])';
           
            fprintf(logfile,'%d trials in muscle ', size(orgData,1));
            fprintf(logfile, '%s\n', muscleName);
            
            newData = [(1:length(orgData))' triggerNumbers orgData];

            % Check that reorganized data has as many rows as events:
            if (length(newData) ~= max(format_data.Event_Number))
               fprintf(logfile,'Error - data size mismatch with event number in subject %d', b);
               fprintf(logfile,' in muscle: %s!\n', muscleName);
            end
           
            % Create baseline ...
            baseStart = 3;
            baseEnd = 3 + num_baseline_per_trial - 1;
            baseData = newData(:,baseStart:baseEnd);
            dataStart = baseEnd + 1;
            trialData = newData(:,dataStart:end);

            switch(baseType)
               case 'trial'
                   % mean per trial
                   baseline = nanmean(baseData,2);
               case 'median'
                   % overall median
                   baseline = nanmedian(reshape(baseData,1,[]));
               case 'window'
                   % sliding window
                   trialBase = nanmean(baseData,2); % first, average per trial
                   baseline = nanmoving_average(trialBase,windowLength,1); % then, smooth using previous/next trial means
               case 'min'
                   minBase = min(baseData,[],2);
                   baseline = nanmoving_average(minBase,windowLength,1); % then, smooth using previous/next trial means                    
               otherwise
                   % assume per trial baseline
                   baseline = nanmean(baseData,2);
            end

            % Remove baseline and finalize trial data ...
            if(isscalar(baseline))
               diffData = trialData - baseline;
            else
               diffData = trialData - repmat(baseline,1,size(trialData,2));
            end

            finalMuscleData = [newData(:,1:2) diffData];

            % Add to muscle data file, which includes all subjects:
            eval([char(muscVarNames{i,1}) '_Data = [' char(muscVarNames{i,1}) '_Data; [b*ones(length(finalMuscleData),1) finalMuscleData]];']);
            
        end
        
        % Log the work that was completed ...
        fprintf(logfile,'Finished working on subject %d!\n',b);

    end

end

fprintf(logfile,'\n### SAVING DATA ###\r');

%% SAVE DATA
for i = 1:numMuscles
        
        muscleSaveName = char(muscleNames{i,1});
        muscleVarName = char(muscVarNames{i,1});
        
        file_name = ['Physio(EMG) Excel data - ' muscleSaveName ' - ' datestr(today)];
        header = ['subject' 'trial' 'trigger' num2cell(intervalLength:intervalLength:trialLength)];
        xlswrite(file_name,[header; num2cell(eval([muscleVarName '_Data']))]);
        
        if (i == 1)
            save(['Physio(EMG) MATLAB Data ' datestr(today)],[muscleVarName '_Data']);
        else
            save(['Physio(EMG) MATLAB Data ' datestr(today)],[muscleVarName '_Data'],'-append');
        end
        
        fprintf(logfile,'Saved data for %s muscle!\n', muscleSaveName);
    
end

% Close logfile:
fclose(logfile);