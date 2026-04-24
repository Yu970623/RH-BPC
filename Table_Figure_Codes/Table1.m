%% Generate LaTeX tables for Section 5.3.1 directly from raw result files
% This script does NOT depend on any intermediate output from other scripts.
% It directly reads raw result files of all algorithms and exports LaTeX tables.
%
% Output:
%   1) Table_531_FinalQuality_BySet.tex
%   2) Table_531_FinalQuality_OverallOnly.tex
%
% Metrics:
%   - Final Sum_obj
%   - Final DroneEnergy(Wh)
%   - Total NewSortieCreated
%   - Final NumSorties
%
% -------------------------------------------------------------------------

clear; clc;

%% ========================= User settings ===============================
rootDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment';
outDir  = fullfile(rootDir, 'Analysis_5_3_1_from_raw');
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

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

allSets  = {Set1, Set2, Set3};
setNames = {'Set1-15', 'Set2-30', 'Set3-50'};
groupOrder = {'Set1-15', 'Set2-30', 'Set3-50', 'Overall'};

metricDefs = {
    'Final_SumObj',         'Final $Sum\_obj$';
    'Final_DroneEnergy_Wh', 'Final energy (Wh)';
    'Total_NewRoutes',      'Total new routes';
    'Final_NumSorties',     'Final sorties'
    };

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
            sourceFlag = "OK";

            if isfile(xlsxFile)
                try
                    T = readtable(xlsxFile, 'VariableNamingRule', 'preserve');
                    loaded = true;
                catch ME
                    warning('Failed to read xlsx: %s\nReason: %s', xlsxFile, ME.message);
                end
            end

            if ~loaded && isfile(csvFile)
                try
                    T = readtable(csvFile, 'VariableNamingRule', 'preserve');
                    loaded = true;
                catch ME
                    warning('Failed to read csv: %s\nReason: %s', csvFile, ME.message);
                end
            end

            if ~loaded
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, "MissingFile"}; %#ok<SAGROW>
                continue;
            end

            if isempty(T)
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, "EmptyTable"}; %#ok<SAGROW>
                continue;
            end

            % ---- Find required columns robustly ----
            colStage      = find_col(T, {'Stage'});
            colStatus     = find_col(T, {'Status'});
            colSumObj     = find_col(T, {'Sum_obj','Sum obj','SumObj'});
            colEnergy     = find_col(T, {'DroneEnergy(Wh)','DroneEnergy','Energy'});
            colNewRoute   = find_col(T, {'NewSortieCreated','NewRouteCreated','New Routes','NewRoutes'});
            colNumSorties = find_col(T, {'NumSorties','Num Sorties','NumSortie'});

            if isempty(colStage) || isempty(colStatus) || isempty(colSumObj) || isempty(colEnergy) || isempty(colNewRoute)
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, "MissingColumn"}; %#ok<SAGROW>
                continue;
            end

            statusVec = string(T.(colStatus));
            validMask = ~ismissing(statusVec) & strlength(strtrim(statusVec)) > 0;
            Tvalid = T(validMask, :);

            if isempty(Tvalid)
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, "NoValidRow"}; %#ok<SAGROW>
                continue;
            end

            % ---- Prefer last OK stage ----
            statusValid = string(Tvalid.(colStatus));
            okMask = strcmpi(strtrim(statusValid), 'OK');

            if any(okMask)
                Tok = Tvalid(okMask, :);
                stageOk = Tok.(colStage);
                [~, idxLast] = max(stageOk);
                Tlast = Tok(idxLast, :);
                finalStatus = "OK";
            else
                stageValid = Tvalid.(colStage);
                [~, idxLast] = max(stageValid);
                Tlast = Tvalid(idxLast, :);
                finalStatus = "NonOKFinal";
            end

            finalSumObj = double(Tlast.(colSumObj));
            finalEnergy = double(Tlast.(colEnergy));

            if ~isempty(colNumSorties)
                finalNumSorties = double(Tlast.(colNumSorties));
            else
                finalNumSorties = NaN;
            end

            % total new routes across all valid stages
            totalNewRoutes = nansum(double(Tvalid.(colNewRoute)));

            rows(end+1, :) = {setNames{s}, instName, alg(a).name, ...
                              finalSumObj, finalEnergy, totalNewRoutes, ...
                              finalNumSorties, finalStatus}; %#ok<SAGROW>
        end
    end
end

resultTbl = cell2table(rows, 'VariableNames', ...
    {'SetName','Instance','Algorithm','Final_SumObj','Final_DroneEnergy_Wh', ...
     'Total_NewRoutes','Final_NumSorties','StatusFlag'});

% Export raw extracted long table for checking
writetable(resultTbl, fullfile(outDir, 'FinalMetrics_Long_FromRaw.xlsx'));
writetable(resultTbl, fullfile(outDir, 'FinalMetrics_Long_FromRaw.csv'));

%% ========================= Build summary ===============================
T = resultTbl(strcmp(resultTbl.StatusFlag, "OK"), :);

sumRows = {};
for g = 1:numel(groupOrder)
    grp = groupOrder{g};

    if strcmp(grp, 'Overall')
        Tg = T;
    else
        Tg = T(strcmp(string(T.SetName), grp), :);
    end

    for m = 1:size(metricDefs,1)
        metricVar = metricDefs{m,1};
        metricLab = metricDefs{m,2};

        row = cell(1, 2 + numel(algOrder));
        row{1} = grp;
        row{2} = metricLab;

        for a = 1:numel(algOrder)
            Ta = Tg(strcmp(string(Tg.Algorithm), algOrder{a}), :);

            if isempty(Ta) || ~ismember(metricVar, Ta.Properties.VariableNames)
                row{2+a} = '--';
            else
                x = Ta.(metricVar);
                x = x(~isnan(x));
                if isempty(x)
                    row{2+a} = '--';
                else
                    row{2+a} = format_mean_std(x, metricVar);
                end
            end
        end

        sumRows(end+1, :) = row; %#ok<SAGROW>
    end
end

sumTbl = cell2table(sumRows, 'VariableNames', ...
    [{'Group','Metric'}, matlab.lang.makeValidName(algOrder)]);

writetable(sumTbl, fullfile(outDir, 'FinalMetrics_Summary_FromRaw.xlsx'));
writetable(sumTbl, fullfile(outDir, 'FinalMetrics_Summary_FromRaw.csv'));

%% ========================= Write LaTeX table: By Set ===================
outFile1 = fullfile(outDir, 'Table_531_FinalQuality_BySet.tex');
fid = fopen(outFile1, 'w', 'n', 'UTF-8');
if fid < 0
    error('Cannot open output file: %s', outFile1);
end

fprintf(fid, '%% Auto-generated LaTeX table for Section 5.3.1 from raw files\n');
fprintf(fid, '\\begin{table*}[t]\n');
fprintf(fid, '\\centering\n');
fprintf(fid, '\\caption{Overall solution quality at the final stage across different instance sets.}\n');
fprintf(fid, '\\label{tab:531_final_quality_byset}\n');
fprintf(fid, '\\scriptsize\n');
fprintf(fid, '\\setlength{\\tabcolsep}{5pt}\n');
fprintf(fid, '\\renewcommand{\\arraystretch}{1.4}\n');
fprintf(fid, '\\begin{tabular}{llcccc}\n');
fprintf(fid, '\\toprule\n');
fprintf(fid, 'Group & Metric & RH-BPC & Greedy & MIP & Repair \\\\\n');
fprintf(fid, '\\midrule\n');

for g = 1:numel(groupOrder)
    grp = groupOrder{g};
    idx = strcmp(sumTbl.Group, grp);
    subT = sumTbl(idx, :);

    for r = 1:height(subT)
        if r == 1
            fprintf(fid, '\\multirow{4}{*}{%s} & %s & %s & %s & %s & %s \\\\\n', ...
                latex_escape(grp), ...
                subT.Metric{r}, ...
                latex_escape(subT.RH_BPC{r}), ...
                latex_escape(subT.Greedy{r}), ...
                latex_escape(subT.MIP{r}), ...
                latex_escape(subT.Repair{r}));
        else
            fprintf(fid, ' & %s & %s & %s & %s & %s \\\\\n', ...
                subT.Metric{r}, ...
                latex_escape(subT.RH_BPC{r}), ...
                latex_escape(subT.Greedy{r}), ...
                latex_escape(subT.MIP{r}), ...
                latex_escape(subT.Repair{r}));
        end
    end

    if g < numel(groupOrder)
        fprintf(fid, '\\midrule\n');
    end
end

fprintf(fid, '\\bottomrule\n');
fprintf(fid, '\\end{tabular}\n');
fprintf(fid, '\\end{table*}\n');
fclose(fid);

%% ========================= Write LaTeX table: Overall ==================
outFile2 = fullfile(outDir, 'Table_531_FinalQuality_OverallOnly.tex');
fid = fopen(outFile2, 'w', 'n', 'UTF-8');
if fid < 0
    error('Cannot open output file: %s', outFile2);
end

idxOverall = strcmp(sumTbl.Group, 'Overall');
To = sumTbl(idxOverall, :);

fprintf(fid, '%% Auto-generated compact LaTeX table for Section 5.3.1 from raw files\n');
fprintf(fid, '\\begin{table}[t]\n');
fprintf(fid, '\\centering\n');
fprintf(fid, '\\caption{Overall final-stage performance comparison.}\n');
fprintf(fid, '\\label{tab:531_final_quality_overall}\n');
fprintf(fid, '\\scriptsize\n');
fprintf(fid, '\\setlength{\\tabcolsep}{4pt}\n');
fprintf(fid, '\\renewcommand{\\arraystretch}{1.3}\n');
fprintf(fid, '\\begin{tabular}{lcccc}\n');
fprintf(fid, '\\toprule\n');
fprintf(fid, 'Metric & RH-BPC & Greedy & MIP & Repair \\\\\n');
fprintf(fid, '\\midrule\n');

for r = 1:height(To)
    fprintf(fid, '%s & %s & %s & %s & %s \\\\\n', ...
        To.Metric{r}, ...
        latex_escape(To.RH_BPC{r}), ...
        latex_escape(To.Greedy{r}), ...
        latex_escape(To.MIP{r}), ...
        latex_escape(To.Repair{r}));
end

fprintf(fid, '\\bottomrule\n');
fprintf(fid, '\\end{tabular}\n');
fprintf(fid, '\\end{table}\n');
fclose(fid);

disp('------------------------------------------------------------');
disp('Done. Files generated:');
disp(fullfile(outDir, 'FinalMetrics_Long_FromRaw.xlsx'));
disp(fullfile(outDir, 'FinalMetrics_Summary_FromRaw.xlsx'));
disp(outFile1);
disp(outFile2);
disp('------------------------------------------------------------');

%% ========================= Local functions =============================
function colName = find_col(T, candidateNames)
% Find table variable name robustly by exact/normalized matching.
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

    % fallback: contains matching
    for i = 1:numel(candNorm)
        idx = find(contains(varsNorm, candNorm{i}), 1, 'first');
        if ~isempty(idx)
            colName = vars{idx};
            return;
        end
    end
end

function out = normalize_names(in)
% Normalize names by removing blanks and punctuation.
    if ischar(in)
        in = {in};
    end
    out = cell(size(in));
    for k = 1:numel(in)
        s = lower(string(in{k}));
        s = regexprep(s, '[\s_\-\(\)\[\]\{\},./\\]', '');
        out{k} = char(s);
    end
end

function s = format_mean_std(x, metricVar)
% Format mean ± std depending on metric type.
    mu = mean(x, 'omitnan');
    sd = std(x, 'omitnan');

    if contains(metricVar, 'NewRoutes') || contains(metricVar, 'NumSorties')
        s = sprintf('%.2f $\\pm$ %.2f', mu, sd);
    elseif contains(metricVar, 'Energy')
        s = sprintf('%.1f $\\pm$ %.1f', mu, sd);
    else
        s = sprintf('%.3f $\\pm$ %.3f', mu, sd);
    end
end

function s = latex_escape(strIn)
% Escape special LaTeX characters
    s = string(strIn);
    s = replace(s, '\', '\\textbackslash ');
    s = replace(s, '_', '\_');
    s = replace(s, '%', '\%%');
    s = replace(s, '&', '\&');
    s = replace(s, '#', '\#');
    s = replace(s, '{', '\{');
    s = replace(s, '}', '\}');
    s = char(s);
end