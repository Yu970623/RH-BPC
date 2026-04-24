%% Generate strict-format Excel tables for Section 5.3
% This script directly reads the raw result files of RH-BPC / Greedy / MIP / Repair
% and outputs two Excel sheets that strictly follow the user-specified table formats.
%
% Output workbook:
%   Final_Tables_StrictFormat.xlsx
%
% Sheets:
%   Raw_ObjectiveData
%   Table1_FinalComparison
%   Table2_DynamicObjective
%   README

clear; clc;

%% ========================= User settings ===============================
rootDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment';
outDir  = fullfile(rootDir, 'Analysis_FinalTables_StrictFormat');
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

outXlsx = fullfile(outDir, 'Final_Tables_StrictFormat.xlsx');

tieTol = 1e-6;

% ---- Algorithm folders and filename suffixes ----
alg(1).name   = 'RH-BPC';
alg(1).folder = fullfile(rootDir, 'RH-BCP', 'DynamicPickup_details');
alg(1).suffix = '_RH-BPC.xlsx';

alg(2).name   = 'Greedy';
alg(2).folder = fullfile(rootDir, 'Greedy', 'DynamicPickup_details_myopic');
alg(2).suffix = '_Greedy.xlsx';

alg(3).name   = 'MIP';
alg(3).folder = fullfile(rootDir, 'MIP', 'DynamicPickup_details');
alg(3).suffix = '_MIP.xlsx';

alg(4).name   = 'Repair';
alg(4).folder = fullfile(rootDir, 'Repair', 'DynamicPickup_details_repair');
alg(4).suffix = '_Repair.xlsx';

algOrder = {alg.name};

% ---- Instance sets ----
Set1 = {'Ca1-2,3,15', 'Ca1-3,5,15', 'Ca1-6,4,15', 'Ca2-2,3,15', 'Ca2-3,5,15', 'Ca2-6,4,15', ...
        'Ca3-2,3,15', 'Ca3-3,5,15', 'Ca3-6,4,15', 'Ca4-2,3,15', 'Ca4-3,5,15', 'Ca4-6,4,15', ...
        'Ca5-2,3,15', 'Ca5-3,5,15', 'Ca5-6,4,15'};

Set2 = {'Ca1-2,3,30', 'Ca1-3,5,30', 'Ca1-6,4,30', 'Ca2-2,3,30', 'Ca2-3,5,30', 'Ca2-6,4,30', ...
        'Ca3-2,3,30', 'Ca3-3,5,30', 'Ca3-6,4,30', 'Ca4-2,3,30', 'Ca4-3,5,30', 'Ca4-6,4,30', ...
        'Ca5-2,3,30', 'Ca5-3,5,30', 'Ca5-6,4,30'};

Set3 = {'Ca1-2,3,50', 'Ca1-3,5,50', 'Ca1-6,4,50', 'Ca2-2,3,50', 'Ca2-3,5,50', 'Ca2-6,4,50', ...
        'Ca3-2,3,50', 'Ca3-3,5,50', 'Ca3-6,4,50', 'Ca4-2,3,50', 'Ca4-3,5,50', 'Ca4-6,4,50', ...
        'Ca5-2,3,50', 'Ca5-3,5,50', 'Ca5-6,4,50'};

allSets = {Set1, Set2, Set3};
scaleNames = {'15 Customers', '30 Customers', '50 Customers'};
scaleN = [15, 15, 15];

%% ========================= Read raw files ==============================
rows = {};

for s = 1:numel(allSets)
    curSet = allSets{s};

    for i = 1:numel(curSet)
        instName = curSet{i};

        for a = 1:numel(alg)
            xlsxFile = fullfile(alg(a).folder, ['Dynamic_' instName alg(a).suffix]);
            csvFile  = strrep(xlsxFile, '.xlsx', '.csv');

            T = table();
            loaded = false;
            statusFlag = "OK";

            if isfile(xlsxFile)
                try
                    T = readtable(xlsxFile, 'VariableNamingRule', 'preserve');
                    loaded = true;
                catch
                end
            end

            if ~loaded && isfile(csvFile)
                try
                    T = readtable(csvFile, 'VariableNamingRule', 'preserve');
                    loaded = true;
                catch
                end
            end

            if ~loaded || isempty(T)
                rows(end+1, :) = {scaleNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, NaN, "MissingFile"}; %#ok<SAGROW>
                continue;
            end

            colStage  = find_col(T, {'Stage'});
            colStatus = find_col(T, {'Status'});
            colSumObj = find_col(T, {'Sum_obj','Sum obj','SumObj'});

            if isempty(colStage) || isempty(colStatus) || isempty(colSumObj)
                rows(end+1, :) = {scaleNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, NaN, "MissingColumn"}; %#ok<SAGROW>
                continue;
            end

            statusVec = string(T.(colStatus));
            validMask = ~ismissing(statusVec) & strlength(strtrim(statusVec)) > 0;
            Tvalid = T(validMask, :);

            if isempty(Tvalid)
                rows(end+1, :) = {scaleNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, NaN, "NoValidRow"}; %#ok<SAGROW>
                continue;
            end

            okMask = strcmpi(strtrim(string(Tvalid.(colStatus))), 'OK');
            if any(okMask)
                Tuse = Tvalid(okMask, :);
                statusFlag = "OK";
            else
                Tuse = Tvalid;
                statusFlag = "NonOKFinal";
            end

            Tuse = sortrows(Tuse, colStage);

            stageVals = double(Tuse.(colStage));
            objVals   = double(Tuse.(colSumObj));

            if isempty(stageVals) || isempty(objVals)
                rows(end+1, :) = {scaleNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, NaN, "NoObjective"}; %#ok<SAGROW>
                continue;
            end

            idx0 = find(stageVals == 0, 1, 'first');
            if isempty(idx0)
                idx0 = 1;
            end
            [~, idxF] = max(stageVals);

            initObj  = objVals(idx0);
            finalObj = objVals(idxF);
            deltaObj = finalObj - initObj;

            rows(end+1, :) = {scaleNames{s}, instName, alg(a).name, ...
                initObj, finalObj, deltaObj, min(objVals), max(objVals), statusFlag}; %#ok<SAGROW>
        end
    end
end

objTbl = cell2table(rows, 'VariableNames', ...
    {'InstanceScale','Instance','Algorithm','InitialObj','FinalObj','DeltaObj','MinObjAcrossStages','MaxObjAcrossStages','StatusFlag'});

writetable(objTbl, outXlsx, 'Sheet', 'Raw_ObjectiveData');

%% ========================= Prepare valid data ==========================
validTbl = objTbl(strcmp(objTbl.StatusFlag, "OK"), :);
overallTbl = validTbl;
overallTbl.InstanceScale(:) = {'Overall'};
validAll = [validTbl; overallTbl];

groupOrder = {'15 Customers', '30 Customers', '50 Customers', 'Overall'};
groupN = [15, 15, 15, 45];

%% ========================= Table 1 ====================================
header1 = {'Instance','N','Algorithm','Mean Objective values','Median Objective values', ...
           'Min Objective values','Max Objective values','Best Frequency','Wilcoxon rank sum test'};
table1 = header1;

for g = 1:numel(groupOrder)
    grp = groupOrder{g};
    Tg = validAll(strcmp(validAll.InstanceScale, grp), :);

    % best frequency
    instList = unique(Tg.Instance, 'stable');
    bestCount = zeros(1, numel(algOrder));
    for i = 1:numel(instList)
        Ti = Tg(strcmp(Tg.Instance, instList{i}), {'Algorithm','FinalObj'});
        if isempty(Ti), continue; end
        bestVal = min(Ti.FinalObj);
        isBest = abs(Ti.FinalObj - bestVal) <= tieTol;
        for a = 1:numel(algOrder)
            bestCount(a) = bestCount(a) + sum(strcmp(Ti.Algorithm(isBest), algOrder{a}));
        end
    end

    % pairwise strings RH-BPC vs baselines using FinalObj
    pairStr = containers.Map;
    pairStr('RH-BPC') = 'RH-BPC vs.';
    baselines = {'Greedy','MIP','Repair'};

    Tref = Tg(strcmp(Tg.Algorithm, 'RH-BPC'), {'Instance','FinalObj'});
    Tref.Properties.VariableNames = {'Instance','RH_BPC_FinalObj'};

    for b = 1:numel(baselines)
        Tb = Tg(strcmp(Tg.Algorithm, baselines{b}), {'Instance','FinalObj'});
        Tb.Properties.VariableNames = {'Instance','Baseline_FinalObj'};
        Tpair = innerjoin(Tref, Tb, 'Keys', 'Instance');

        if isempty(Tpair)
            pairStr(baselines{b}) = '--';
        else
            diffv = Tpair.RH_BPC_FinalObj - Tpair.Baseline_FinalObj;
            wins   = sum(diffv < -tieTol);
            ties   = sum(abs(diffv) <= tieTol);
            losses = sum(diffv > tieTol);
            pairStr(baselines{b}) = sprintf('%d/%d/%d', wins, ties, losses);
        end
    end

    for a = 1:numel(algOrder)
        Ta = Tg(strcmp(Tg.Algorithm, algOrder{a}), :);

        if isempty(Ta)
            meanVal = NaN; medianVal = NaN; minVal = NaN; maxVal = NaN;
        else
            meanVal   = mean(Ta.FinalObj, 'omitnan');
            medianVal = median(Ta.FinalObj, 'omitnan');
            minVal    = min(Ta.FinalObj);
            maxVal    = max(Ta.FinalObj);
        end

        if a == 1
            row = {grp, groupN(g), algOrder{a}, numfmt(meanVal,2), numfmt(medianVal,2), ...
                numfmt(minVal,2), numfmt(maxVal,2), bestCount(a), pairStr(algOrder{a})};
        else
            row = {'','', algOrder{a}, numfmt(meanVal,2), numfmt(medianVal,2), ...
                numfmt(minVal,2), numfmt(maxVal,2), bestCount(a), pairStr(algOrder{a})};
        end
        table1(end+1, :) = row; %#ok<SAGROW>
    end
end

writecell(table1, outXlsx, 'Sheet', 'Table1_FinalComparison');

%% ========================= Table 2 ====================================
header2 = {'Instance','N','Algorithm', ...
           sprintf('Mean Objective values\n(Initial stage)'), ...
           sprintf('Mean Objective values\n(Final stage)'), ...
           sprintf('Mean Objective values\n(Delta)'), ...
           sprintf('Wilcoxon rank sum test\n(Initial stage)'), ...
           sprintf('Wilcoxon rank sum test\n(Final stage)'), ...
           sprintf('Wilcoxon rank sum test\n(Delta)')};

table2 = header2;

for g = 1:numel(groupOrder)
    grp = groupOrder{g};
    Tg = validAll(strcmp(validAll.InstanceScale, grp), :);

    pairInit  = containers.Map;
    pairFinal = containers.Map;
    pairDelta = containers.Map;
    pairInit('RH-BPC')  = 'RH-BPC vs.';
    pairFinal('RH-BPC') = 'RH-BPC vs.';
    pairDelta('RH-BPC') = 'RH-BPC vs.';

    baselines = {'Greedy','MIP','Repair'};
    Tref = Tg(strcmp(Tg.Algorithm, 'RH-BPC'), {'Instance','InitialObj','FinalObj','DeltaObj'});
    Tref.Properties.VariableNames = {'Instance','RH_BPC_Initial','RH_BPC_Final','RH_BPC_Delta'};

    for b = 1:numel(baselines)
        Tb = Tg(strcmp(Tg.Algorithm, baselines{b}), {'Instance','InitialObj','FinalObj','DeltaObj'});
        Tb.Properties.VariableNames = {'Instance','B_Initial','B_Final','B_Delta'};
        Tpair = innerjoin(Tref, Tb, 'Keys', 'Instance');

        if isempty(Tpair)
            pairInit(baselines{b})  = '--';
            pairFinal(baselines{b}) = '--';
            pairDelta(baselines{b}) = '--';
        else
            d1 = Tpair.RH_BPC_Initial - Tpair.B_Initial;
            d2 = Tpair.RH_BPC_Final   - Tpair.B_Final;
            d3 = Tpair.RH_BPC_Delta   - Tpair.B_Delta;

            pairInit(baselines{b})  = sprintf('%d/%d/%d', sum(d1 < -tieTol), sum(abs(d1) <= tieTol), sum(d1 > tieTol));
            pairFinal(baselines{b}) = sprintf('%d/%d/%d', sum(d2 < -tieTol), sum(abs(d2) <= tieTol), sum(d2 > tieTol));
            pairDelta(baselines{b}) = sprintf('%d/%d/%d', sum(d3 < -tieTol), sum(abs(d3) <= tieTol), sum(d3 > tieTol));
        end
    end

    for a = 1:numel(algOrder)
        Ta = Tg(strcmp(Tg.Algorithm, algOrder{a}), :);

        if isempty(Ta)
            meanInit = NaN; meanFinal = NaN; meanDelta = NaN;
        else
            meanInit  = mean(Ta.InitialObj, 'omitnan');
            meanFinal = mean(Ta.FinalObj, 'omitnan');
            meanDelta = mean(Ta.DeltaObj, 'omitnan');
        end

        if a == 1
            row = {grp, groupN(g), algOrder{a}, numfmt(meanInit,2), numfmt(meanFinal,2), numfmt(meanDelta,2), ...
                pairInit(algOrder{a}), pairFinal(algOrder{a}), pairDelta(algOrder{a})};
        else
            row = {'','', algOrder{a}, numfmt(meanInit,2), numfmt(meanFinal,2), numfmt(meanDelta,2), ...
                pairInit(algOrder{a}), pairFinal(algOrder{a}), pairDelta(algOrder{a})};
        end
        table2(end+1, :) = row; %#ok<SAGROW>
    end
end

writecell(table2, outXlsx, 'Sheet', 'Table2_DynamicObjective');

%% ========================= README =====================================
readme = {
    'Sheet','Description';
    'Raw_ObjectiveData','Instance-level objective data extracted from raw result files';
    'Table1_FinalComparison','Strict-format final-stage comparison table';
    'Table2_DynamicObjective','Strict-format dynamic objective table'
    };
writecell(readme, outXlsx, 'Sheet', 'README');

fprintf('Done. Output workbook:\n%s\n', outXlsx);

%% ========================= Local functions ============================
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

function s = numfmt(x, nd)
    if isnan(x)
        s = '';
    else
        fmt = ['%0.', num2str(nd), 'f'];
        s = sprintf(fmt, x);
    end
end
