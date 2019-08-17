%% OC3 DATACLEANER
%% V1.0.0.1 - GDJ

clear all;
clc;
addpath('utils');
addpath('utils/xlwrite');
addpath('utils/xlwrite/poi_library');
if ismac
    run('utils/xlwrite/firstlaunch_xlwrite.m');
    disp('Added Mac export compatibility');
end

[file, path] = uigetfile('*.xlsx');
cd(path);

[dataInt, dataStr, dataRaw] = xlsread(fullfile(path,file));

studyEventDefinitionLibrary = {'Study Event Definition'};
crfLibrary = {'CoreProtocol_PD', 'Behavioural_CoreProtocol', 'CoreProtocol_FollowUp_PD'};
disp('Working...');
%%% METADATA ANALYSIS
%% Remove NaNs from Raw Data
% dataRaw(cellfun(@(dataRaw) any(isnan(dataRaw)),dataRaw)) = [];
numRows = size(dataRaw,1);
numCols = size(dataRaw,2);
for currentRow = 1:numRows
    for currentCol = 1:numCols
        if isnan(dataRaw{currentRow,currentCol})
            dataRaw{currentRow,currentCol} = [];
        end
    end
end

%% 1) Find Study Metadata
% Find empty line of rows to delineate between metadata and data
numRows = size(dataRaw,1);
numCols = size(dataRaw,2);
rowCounter = 1;

emptyRows = [];
studyMetadata = [];
studyData = [];

for currentRow = 1:numRows
    if isempty(dataRaw{currentRow,1})
        for currentColumn = 1:numCols
            if isempty(dataRaw{currentRow,currentColumn})
                if currentColumn == numCols
                    emptyRows(rowCounter) = currentRow;
                    rowCounter = rowCounter + 1;
                end
            else
                break;
            end
        end
    end
end
disp('Found metadata');

% emptyRows = emptyRows(diff(emptyRows)==1);

% Create Metadata structure
studyMetadata = dataStr(1:max(emptyRows)-1,:);

% Remove MetaData from dataRaw
studyData = dataRaw(max(emptyRows)+1:end,:);
studyDataStr = dataStr(max(emptyRows)+1:end,:);

% Find Study Definitions and Indices
numRows = size(studyMetadata,1);
numCols = size(studyMetadata,2);

numStudyEvents = 0;
studyEvents = {};
eventNames = {};

for currentRow = 1:numRows
    if contains(studyMetadata{currentRow,1},studyEventDefinitionLibrary)
        for currentCol = 1:numCols
            if ~contains(studyMetadata{currentRow,currentCol},studyEventDefinitionLibrary) && contains(studyMetadata{currentRow,currentCol},'E')
                numStudyEvents = numStudyEvents + 1;
                studyEvents = [studyEvents; {studyMetadata{currentRow,currentCol},currentRow,currentCol}];
                eventNames{numStudyEvents} = studyMetadata{currentRow,currentCol-1};
                disp(['Found event: ''', eventNames{numStudyEvents}, '''']);
                eventNames = strrep(eventNames,' ','_');
            end
        end
    end
end

%% Find Associated CRFs for each Event
studyCRF = {};
searchArea = studyMetadata(studyEvents{1,2}:studyEvents{end,2},2:3);
if size(searchArea,1)==1
    searchArea = studyMetadata(studyEvents{1,2}:end,2:3);
end

% Remove empties because why not?
for currentRow = size(searchArea,1):-1:1
    if isempty(searchArea{currentRow,1})
        searchArea(currentRow,:) = [];
    end
end

% Find CRF Versions
crfVersions = {};
versionCounter = 1;
if size(searchArea,1) >1
    for currentRow = 1:size(searchArea,1)
        [startIndex, endIndex] = regexp(searchArea{currentRow,1},' - ');
        if ~isempty(startIndex) && ~isempty(endIndex)
            crfVersions(versionCounter,:) = searchArea(currentRow,:);
            versionCounter = versionCounter + 1;
        end
    end
else
end

for currentRow = 1:size(crfVersions,1)
    % Remove trailing version numbers from CRF names
    [startIndex, endIndex] = regexp(crfVersions{currentRow,1},' - ');
    if ~isempty(startIndex) && ~isempty(endIndex)
        crfVersions{currentRow,1} = crfVersions{currentRow,1}(1:startIndex-1);
    end
end

% Sort crfVersions array
crfVersions = sort(crfVersions);

% Get names of each CRF
[crfNames, crfIndices, vals] = unique(crfVersions(:,1));

versionArray = cell(length(crfNames),length(vals(vals==(mode(vals))))); % Create empty cell array
versionArray(:,:) = {NaN};

for currentCRF = 1:size(crfNames,1)
    if currentCRF == size(crfNames,1)
        versionArray(currentCRF,1:length(crfVersions(crfIndices(currentCRF):end,2)')) = crfVersions(crfIndices(currentCRF):end,2)';
    else
        versionArray(currentCRF,1:length(crfVersions(crfIndices(currentCRF):crfIndices(currentCRF+1)-1,2))) = crfVersions(crfIndices(currentCRF):crfIndices(currentCRF+1)-1,2);
    end
end

%% SPLIT INTO 3D CELL ARRAY OF EVENTS
eventMatrix = {};
currentRow = 1;
eventCounter = ones(1,size(studyEvents,1));
for currentColumn = 1:size(studyData,2)
    str = studyData{currentRow,currentColumn};
    if contains(str,studyEvents(:,1))
        for currentEvent = 1:size(studyEvents(:,1),1)
            if contains(str,studyEvents(currentEvent,1))
                eventMatrix(:,eventCounter(currentEvent),currentEvent) = studyData(:,currentColumn);
                eventCounter(currentEvent) = eventCounter(currentEvent) + 1;
            end
        end
    end
end

% Create participant data matrix of data not belonging to Event Codes
b = find(contains(studyData(1,:),'_E'),1,'first');
genData = studyData(:,1:b-1);

%% FIND DUPLICATE CELLS

% append underscores to studyEvents for more accurate filtering
for i = 1:size(studyEvents,1)
    studyEvents{i,1} = ['_', studyEvents{i,1}, '_'];
end

crfs = {};
disp('Tidying duplicate parameter headers...');
for currentEvent = 1:size(eventMatrix,3)
    headers = eventMatrix(:,:,currentEvent);
    % Find all mentions of the duplicate
    for currentCRFArray = 1:size(studyEvents,1)
        crfs = versionArray(currentCRFArray,:);
        crfs(cellfun(@(crfs) any(isnan(crfs)),crfs)) = []; % remove NaNs
        
        % Remove trailing information for comparison of strings
        expression = studyEvents{currentCRFArray,1};
        [startIndices, endIndices] = regexp(headers(1,:),expression);
        
        for i = 1:size(headers,2)
            if ~isempty(endIndices{i})
                headers{1,i} = headers{1,i}(1:startIndices{i}-1);
            end
        end
    end
    eventMatrix(:,:,currentEvent) = headers;
end
disp('Done');

clear tempHeaders headers

% remove trailing whitespace
eventMatrix(1,:,:) = strtrim(eventMatrix(1,:,:));

% find unique strings in header

clear newMatrix;
newMatrix = cell(size(eventMatrix));
for currentEvent = 1:size(eventMatrix,3)
    colCounter = 1;
    disp('--------------------------------');
    headers = eventMatrix(1,:,currentEvent);
    headers(cellfun(@(tempUniqueStrings) any(isempty(tempUniqueStrings)),headers)) = []; % Remove empty cells
    headers(cellfun(@(tempUniqueStrings) any(isnan(tempUniqueStrings)),headers)) = []; % Remove cells with NaNs
    uniqueHeaders = unique(headers,'stable'); % Unique header elements
    
    disp(['Found ', num2str(numel(uniqueHeaders)), ' unique headers for event ', num2str(currentEvent)]);
    
    for currentParamString = 1:size(uniqueHeaders,2) % Loop through each unique column header
        clear singleParameterIdx duplicateParameters unifiedParameter force
        
        singleParameterIdx = ismember(headers, uniqueHeaders{currentParamString});
        paramDuplicates = eventMatrix(:,singleParameterIdx,currentEvent);
        
        for currentDuplicate = 2:size(paramDuplicates,1)
            nonEmptyIdx = find(~cellfun(@isempty,paramDuplicates(currentDuplicate,:)));
            if ~isempty(nonEmptyIdx)
                if numel(paramDuplicates(currentDuplicate,:)) > 1
                    if isequal(paramDuplicates,paramDuplicates)
                        unifiedParameter(currentDuplicate) = paramDuplicates(currentDuplicate,nonEmptyIdx(1));
                    else
                        error('ERROR: Non-identical duplicate data discovered for participant');
                    end
                elseif numel(paramDuplicates(currentDuplicate,:)) == 1
                    unifiedParameter(currentDuplicate) = paramDuplicates(currentDuplicate,nonEmptyIdx);
                end
            else
                unifiedParameter{currentDuplicate} = [];
            end
        end
        
        unifiedParameter(1) = paramDuplicates(1,1);
        newMatrix(:,colCounter,currentEvent) = unifiedParameter';
        colCounter = colCounter + 1;
        clear singleParameterIdx duplicateParameters unifiedParameter
    end
    disp([num2str(colCounter - 1), ' Parameters processed']);
    
    %     end
end

%% Find Specific Questionnaires
% TOTHINK: A BETTER APPROACH WOULD BE DOWN THE TWO-TIER "UNIQUE-NON-UNIQUE"
% PATHWAY - SUBSTRING 1 SHOULD ALWAYS BE UNIQUE, SUBSTRING 2 SHOULD NEVER
% BE UNIQUE...
global selectedItems;
selectSubs = false;
selectSubs = questdlg('Select specific questionnaires/categories?', ...
    'Select Questionnaires', ...
    'Yes','No','Yes');
switch selectSubs
    case 'Yes'
        selectSubs = true;
    case 'No'
        selectSubs = false;
end

if selectSubs
    
    questionnaireNames = {};
    questionnaireHeaders = permute(newMatrix(1,:,:),[2 1 3]); % Define headers
    questionnaireHeaders = questionnaireHeaders(:);
    
    emptyCells = find(cellfun(@isempty,questionnaireHeaders)); % find the empty cells in the headers
    questionnaireHeaders(emptyCells) = []; %remove the empty cells
    expression = '_';
    words = regexp(questionnaireHeaders,expression,'split'); % [cellSizesRow, cellSizesCols] = cell2mat(cellfun(@size,words,'UniformOutput',false)); %TOTHINK: Is there any value in using the size vector over the length scalar?
    cellLengths = cellfun(@length,words); % find the length of each split header
    words = words(cellLengths(:)>1); % remove single string cells (as won't contain an _)
    
    for i = 1:size(words,1)
        questionnaireNames{i,1} = words{i}{1,1};
        if length(words{i}) == 2
            questionnaireNames{i,2} = words{i}{1,2};
        end
    end
    uniqueQuestionnaires = unique(questionnaireNames(:,1));

    fig = uifigure('Position',[0,0,200,350],'Visible','off');
    movegui(fig,'center');
    fig.Visible = 'on';
    uilabel(fig,'Text','Select Categories:','Position',[10,310,180,40],'FontWeight','normal');
    catList = uilistbox(fig,'Position',[10 40 180 280],'Items',uniqueQuestionnaires,'ItemsData',uniqueQuestionnaires,'Multiselect','on');
    btnDone = uibutton(fig,'Position',[10 10 180 25],'Text','Done','ButtonPushedFcn', @(btn,event) getListBoxItems(btn,catList.Value));
    
    while isempty(selectedItems)
        drawnow
    end
    
    % Now remove all columns from the array that don't contain the selected
    % items
    
    for currentLeaf = 1:size(newMatrix,3)
        tempTopRow = newMatrix(1,:,currentLeaf);
        tempTopRow(cellfun(@(tempTopRow) any(isempty(tempTopRow)),tempTopRow)) = []; % Remove empty cells
        tempTopRow(cellfun(@(tempTopRow) any(isnan(tempTopRow)),tempTopRow)) = []; % Remove cells with NaNs
        selectedItems
        keepIdx = contains(tempTopRow,selectedItems);
        tempMatrix = newMatrix(:,keepIdx==1,currentLeaf);
        newMatrix(:,:,currentLeaf) = cell(size(newMatrix,1),size(newMatrix,2));
        newMatrix(:,1:numel(keepIdx(keepIdx==1)),currentLeaf) = tempMatrix;
    end

end

%% Add General Participant Data (not belonging to events) back to each leaf in 3D matrix
for currentLeaf = 1:size(newMatrix,3)
    newMatrix(:,1:end+size(genData,2),currentLeaf) = [genData, newMatrix(:,:,currentLeaf)];
end

%% Add Patient Diagnoses
addDiagnosisQ = questdlg('Attach diagnoses to each participant?', ...
    'Add diagnoses', ...
    'Yes','No','Yes');
switch addDiagnosisQ
    case 'Yes'
        newMatrix = appendDiagnoses(newMatrix);
    case 'No'
        disp('No diagnoses appended.');
end

%% Prepare For Export
% Sort rows in descending order of most populated columns and then export
warning off;
filename = [fullfile(path),'CLEANED_',num2str(floor(posixtime(datetime('now')))),'_', fullfile(file)];
disp('------------------------------');
if ismac
    run firstlaunch_xlwrite.m
    xlwrite(filename,studyMetadata,'Metadata');
    disp('Exported Metadata');
else
    xlswrite(filename,studyMetadata,'Metadata');
    disp('Exported Metadata');
end
disp('------------------------------');


for currentEvent = 1:size(newMatrix,3)
    numFilledColumns = sum(~cellfun(@isempty,newMatrix(:,:,currentEvent)),2);
    [ranks_ordered, idx] = sort(numFilledColumns, 'descend');
    newMatrix(:,:,currentEvent) = newMatrix(idx,:,currentEvent);
    sheet = eventNames{currentEvent};
    if ismac
        xlwrite(filename,newMatrix(:,:,currentEvent),sheet);
    else
        xlswrite(filename,newMatrix(:,:,currentEvent),sheet);
    end
    disp(['Exported ', eventNames{currentEvent}]);
    disp('------------------------------');
end
warning on;

if ispc
    sheetName = 'Sheet';
    objExcel = actxserver('Excel.Application');
    objExcel.Workbooks.Open(filename); % Full path is necessary!
    try
        % Throws an error if the sheets do not exist.
        objExcel.ActiveWorkbook.Worksheets.Item([sheetName '1']).Delete;
        objExcel.ActiveWorkbook.Worksheets.Item([sheetName '2']).Delete;
        objExcel.ActiveWorkbook.Worksheets.Item([sheetName '3']).Delete;
    catch
        
    end
    % Save, close and clean up.
    objExcel.ActiveWorkbook.Save;
    objExcel.ActiveWorkbook.Close;
    objExcel.Quit;
    objExcel.delete;
    
    % For MatLab 2019A
    % writecell(newMatrix,'test.xlsx','Sheet','A1');
end

data_cleaned = newMatrix;
disp('------------------------------');
fprintf('Done.\nCleaned data saved to:');
disp(filename);
disp('Data are available in the MatLab workspace as ''data_cleaned''');

try
    winopen(path);
catch
    system(['open ',path]);
end

clearvars -except data_cleaned

function [inputData] = appendDiagnoses(inputData)
        % Prompt for list of diagnoses
        if ~strcmp(inputData{1, 2, 1},'Diagnosis') %Add diagnosis column if not already exists
            inputData = [inputData(:,1,:),cell(size(inputData(:,1),1),1,size(inputData,3)),inputData(:,2:end,:)];
            inputData(1,2,:) = {'Diagnosis'};
        end
        
        [dFile, dPath] = uigetfile('*xlsx','Select Clinical Conductor patient list'); % get list of CC diagnoses
        [ddInt, ddStr, ddRaw] = xlsread(fullfile(dPath,dFile));
        startIdx = find(contains(ddStr(:,1),'Participant Code'),1) + 1;
        endIdx = size(ddStr,1);
        participantList = ddStr(startIdx:endIdx,1);
        
        % Prompt for Diagnosis Name
        prompt = {'Enter diagnosis:'};
        dlgtitle = 'Input';
        dims = [1 35];
        definput = {'Diagnosis','hsv'};
        answer = inputdlg(prompt,dlgtitle,dims,definput);
        
        for currentLeaf = 1:size(inputData,3)
            for currentRow = 1:size(inputData,1)
                for currentParticipant = 1:size(participantList,1)
                    if strcmp(inputData{currentRow,1,currentLeaf},participantList{currentParticipant,1})
                        inputData{currentRow,2,currentLeaf} = char(answer);
                        disp(['Appended diagnosis to participant ', inputData{currentRow,1,currentLeaf}]);
                    end
                end
            end
        end
        
        addAnotherDiagnosisQ = questdlg('Append more diagnoses?', ...
            'Add diagnoses', ...
            'Yes','No','Yes');
        switch addAnotherDiagnosisQ
            case 'Yes'
                inputData = appendDiagnoses(inputData);
            case 'No'

        end
end

function selectedItems = getListBoxItems(btn,selectedItemsLBox)
global selectedItems;
disp('Items selected:');
selectedItems = selectedItemsLBox'
close gcf force
end
