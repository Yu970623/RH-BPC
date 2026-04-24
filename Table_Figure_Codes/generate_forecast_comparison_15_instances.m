%% Compare RH-BPC, RH-BPC-Forecast (v3), and RH-BPC-Perfect Forecast (v2)
% This script extracts FINAL-STAGE metrics for the 15-customer instances only:
%   - NumSorties
%   - Sum_obj
%   - DroneEnergy(Wh)
%
% Methods compared:
%   1) RH-BPC                        -> online, no future information
%   2) RH-BPC-Forecast              -> forecast-aware dynamic replanning (v3)
%   3) RH-BPC-Perfect Forecast      -> full-information benchmark (v2)
%
% Output format follows the user-requested table structure:
%
% Instances | NumSorties (3 cols) | Sum_obj (3 cols) | DroneEnergy(Wh) (3 cols)
%
% Output files:
%   - PerfectForecast_15Customer_Comparison.xlsx
%   - PerfectForecast_15Customer_Comparison.csv
%
% Notes:
%   - Final-stage row = row with maximum Stage among valid rows
%   - If Status exists, rows with Status='OK' are preferred
%   - This script is intended for the subsection:
%       "Value of perfect forecast information in dynamic replanning"

clear; clc;

%% ========================= User settings ===============================
baseDir1 = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\RH-BCP\DynamicPickup_details';
baseDir2 = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\RH-BCP v2\DynamicPickup_details';
baseDir3 = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\RH-BCP v3\DynamicPickup_details';

outDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\Forecast_Comparison_15';
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

outXlsx = fullfile(outDir, 'PerfectForecast_15Customer_Comparison.xlsx');
outCsv  = fullfile(outDir, 'PerfectForecast_15Customer_Comparison.csv');

% method display names and file suffixes
methods(1).name   = 'RH-BPC';
methods(1).folder = baseDir1;
methods(1).suffix = '_RH-BPC.xlsx';

methods(2).name   = 'RH-BPC-Forecast';
methods(2).folder = baseDir3;
methods(2).suffix = '.xlsx';   % adjust if your v3 filenames have a suffix

methods(3).name   = 'RH-BPC-Perfect Forecast';
methods(3).folder = baseDir2;
methods(3).suffix = '.xlsx';   % adjust if your v2 filenames have a suffix

instances = {'Ca1-2,3,15', 'Ca1-3,5,15', 'Ca1-6,4,15', ...
             'Ca2-2,3,15', 'Ca2-3,5,15', 'Ca2-6,4,15', ...
             'Ca3-2,3,15', 'Ca3-3,5,15', 'Ca3-6,4,15', ...
             'Ca4-2,3,15', 'Ca4-3,5,15', 'Ca4-6,4,15', ...
             'Ca5-2,3,15', 'Ca5-3,5,15', 'Ca5-6,4,15'};

%% ========================= Extract final-stage data ====================
rawRows = {};
tableRows = {};

for i = 1:numel(instances)
    inst = instances{i};

    numSortiesVals = cell(1, numel(methods));
    sumObjVals     = cell(1, numel(methods));
    energyVals     = cell(1, numel(methods));

    for m = 1:numel(methods)
        [finalNumSorties, finalSumObj, finalEnergy, statusFlag, fileUsed] = ...
            read_final_stage_metrics(methods(m).folder, inst, methods(m).suffix);

        numSortiesVals{m} = num_to_str(finalNumSorties, 2);
        sumObjVals{m}     = num_to_str(finalSumObj, 2);
        energyVals{m}     = num_to_str(finalEnergy, 3);

        rawRows(end+1,:) = {inst, methods(m).name, finalNumSorties, finalSumObj, finalEnergy, statusFlag, fileUsed}; %#ok<SAGROW>
    end

    row = [{inst}, numSortiesVals, sumObjVals, energyVals];
    tableRows(end+1,:) = row; %#ok<SAGROW>
end

%% ========================= Build structured output =====================
header = {'Instances', ...
          'NumSorties_RH-BPC', 'NumSorties_RH-BPC-Forecast', 'NumSorties_RH-BPC-PerfectForecast', ...
          'Sum_obj_RH-BPC', 'Sum_obj_RH-BPC-Forecast', 'Sum_obj_RH-BPC-PerfectForecast', ...
          'DroneEnergyWh_RH-BPC', 'DroneEnergyWh_RH-BPC-Forecast', 'DroneEnergyWh_RH-BPC-PerfectForecast'};

structured = [header; tableRows];

writecell(structured, outXlsx, 'Sheet', 'Comparison_Table');

rawTbl = cell2table(rawRows, 'VariableNames', ...
    {'Instance','Method','FinalNumSorties','FinalSumObj','FinalDroneEnergyWh','StatusFlag','FileUsed'});
writetable(rawTbl, outXlsx, 'Sheet', 'Raw_FinalStage_Data');

structuredTbl = cell2table(tableRows, 'VariableNames', matlab.lang.makeValidName(header));
writetable(structuredTbl, outCsv);

disp('============================================================');
disp('Done. Files generated:');
disp(outXlsx);
disp(outCsv);
disp('============================================================');

%% ========================= Local functions ============================
function [finalNumSorties, finalSumObj, finalEnergy, statusFlag, fileUsed] = ...
    read_final_stage_metrics(folderPath, instName, suffix)

    finalNumSorties = NaN;
    finalSumObj     = NaN;
    finalEnergy     = NaN;
    statusFlag      = "MissingFile";
    fileUsed        = "";

    % Candidate filenames:
    % 1) Dynamic_<inst><suffix>
    % 2) Dynamic_<inst>.xlsx
    % 3) Any xlsx whose name contains <inst>
    cand1 = fullfile(folderPath, ['Dynamic_' instName suffix]);
    cand2 = fullfile(folderPath, ['Dynamic_' instName '.xlsx']);
    cand3 = fullfile(folderPath, ['Dynamic_' instName '.csv']);

    filePath = "";

    if isfile(cand1)
        filePath = string(cand1);
    elseif isfile(cand2)
        filePath = string(cand2);
    elseif isfile(cand3)
        filePath = string(cand3);
    else
        D = dir(fullfile(folderPath, ['*' instName '*.xlsx']));
        if ~isempty(D)
            filePath = string(fullfile(folderPath, D(1).name));
        else
            D = dir(fullfile(folderPath, ['*' instName '*.csv']));
            if ~isempty(D)
                filePath = string(fullfile(folderPath, D(1).name));
            end
        end
    end

    if strlength(filePath) == 0
        return;
    end

    fileUsed = char(filePath);

    T = table();
    try
        T = readtable(fileUsed, 'VariableNamingRule', 'preserve');
    catch
        statusFlag = "ReadError";
        return;
    end

    if isempty(T)
        statusFlag = "EmptyFile";
        return;
    end

    colStage  = find_col(T, {'Stage'});
    colSort   = find_col(T, {'NumSorties','Num Sorties','NumSortie'});
    colObj    = find_col(T, {'Sum_obj','Sum obj','SumObj'});
    colEnergy = find_col(T, {'DroneEnergy(Wh)','DroneEnergy','Energy'});
    colStatus = find_col(T, {'Status'});

    if isempty(colStage) || isempty(colSort) || isempty(colObj) || isempty(colEnergy)
        statusFlag = "MissingColumn";
        return;
    end

    if ~isempty(colStatus)
        statusVec = string(T.(colStatus));
        validMask = ~ismissing(statusVec) & strlength(strtrim(statusVec)) > 0;
        T = T(validMask, :);

        if isempty(T)
            statusFlag = "NoValidRow";
            return;
        end

        okMask = strcmpi(strtrim(string(T.(colStatus))), 'OK');
        if any(okMask)
            T = T(okMask, :);
            statusFlag = "OK";
        else
            statusFlag = "NonOKFinal";
        end
    else
        statusFlag = "NoStatusColumn";
    end

    if isempty(T)
        return;
    end

    T = sortrows(T, colStage);
    stageVals = double(T.(colStage));

    if isempty(stageVals)
        statusFlag = "NoStageData";
        return;
    end

    [~, idxF] = max(stageVals);

    finalNumSorties = double(T.(colSort)(idxF));
    finalSumObj     = double(T.(colObj)(idxF));
    finalEnergy     = double(T.(colEnergy)(idxF));
end

function colName = find_col(T, candidateNames)
    vars = T.Properties.VariableNames;
    varsNorm = normalize_names(vars);
    candNorm = normalize_names(candidateNames);
    colName = '';

    for i = 1:numel(candNorm)
        idx = find(strcmp(varsNorm, candNorm{i}), 1, 'first');
        if ~isempty(idx)
            colName = vars{idx};
            return;
        end
    end

    for i = 1:numel(candNorm)
        idx = find(contains(varsNorm, candNorm{i}), 1, 'first');
        if ~isempty(idx)
            colName = vars{idx};
            return;
        end
    end
end

function out = normalize_names(in)
    if ischar(in), in = {in}; end
    out = cell(size(in));
    for k = 1:numel(in)
        s = lower(string(in{k}));
        s = regexprep(s, '[\s_\-\(\)\[\]\{\},./\\]', '');
        out{k} = char(s);
    end
end

function s = num_to_str(x, nd)
    if isnan(x)
        s = '';
    else
        fmt = ['%0.', num2str(nd), 'f'];
        s = sprintf(fmt, x);
    end
end
