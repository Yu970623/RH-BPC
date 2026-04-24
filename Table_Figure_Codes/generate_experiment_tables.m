%% Generate Tables for Final-stage Comparison and Dynamic Responsiveness
% This script directly reads raw result files of four algorithms and outputs
% Excel tables for:
%   Table 1: Final-stage objective comparison
%   Table 2: Dynamic responsiveness statistics
%
% Output workbook:
%   Experiment_Tables_FinalAndDynamic.xlsx
%
% Notes:
%   - The script prefers .xlsx files and falls back to .csv if needed.
%   - "Win" means RH-BPC has a strictly smaller Final Sum_obj than the baseline
%     on the same instance. "Tie" uses abs(diff) <= tieTol.
%   - Dynamic increments are computed as Final - Stage0 for each instance.
%
% Author: ChatGPT

clear; clc;

%% ========================= User settings ===============================
rootDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment';
outDir  = fullfile(rootDir, 'Analysis_Tables_FinalDynamic');
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

outXlsx = fullfile(outDir, 'Experiment_Tables_FinalAndDynamic.xlsx');

% tolerance for tie judgment in pairwise comparison
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
setNames = {'15 Customers', '30 Customers', '50 Customers'};
groupOrder = [setNames, {'Overall'}];

%% ========================= Read raw files ==============================
rowsFinal = {};
rowsStage = {};

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
                rowsFinal(end+1, :) = {setNames{s}, instName, alg(a).name, ...
                    NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, "MissingFile"}; %#ok<SAGROW>
                continue;
            end

            % find columns robustly
            colStage      = find_col(T, {'Stage'});
            colStatus     = find_col(T, {'Status'});
            colSumObj     = find_col(T, {'Sum_obj','Sum obj','SumObj'});
            colEnergy     = find_col(T, {'DroneEnergy(Wh)','DroneEnergy','Energy'});
            colNewRoute   = find_col(T, {'NewSortieCreated','NewRouteCreated','New Routes','NewRoutes'});
            colNumSorties = find_col(T, {'NumSorties','Num Sorties','NumSortie'});

            if isempty(colStage) || isempty(colStatus) || isempty(colSumObj) || isempty(colEnergy) || isempty(colNewRoute) || isempty(colNumSorties)
                rowsFinal(end+1, :) = {setNames{s}, instName, alg(a).name, ...
                    NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, "MissingColumn"}; %#ok<SAGROW>
                continue;
            end

            statusVec = string(T.(colStatus));
            validMask = ~ismissing(statusVec) & strlength(strtrim(statusVec)) > 0;
            Tvalid = T(validMask, :);

            if isempty(Tvalid)
                rowsFinal(end+1, :) = {setNames{s}, instName, alg(a).name, ...
                    NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, "NoValidRow"}; %#ok<SAGROW>
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

            stageVals    = double(Tuse.(colStage));
            sumVals      = double(Tuse.(colSumObj));
            energyVals   = double(Tuse.(colEnergy));
            routeVals    = double(Tuse.(colNumSorties));
            newRouteVals = double(Tuse.(colNewRoute));

            if isempty(stageVals) || isempty(sumVals)
                rowsFinal(end+1, :) = {setNames{s}, instName, alg(a).name, ...
                    NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, NaN, "NoStageData"}; %#ok<SAGROW>
                continue;
            end

            idx0 = find(stageVals == 0, 1, 'first');
            if isempty(idx0)
                idx0 = 1;
            end
            [~, idxF] = max(stageVals);

            stage0Sum     = sumVals(idx0);
            finalSum      = sumVals(idxF);
            stage0Energy  = energyVals(idx0);
            finalEnergy   = energyVals(idxF);
            stage0Route   = routeVals(idx0);
            finalRoute    = routeVals(idxF);

            deltaSum      = finalSum - stage0Sum;
            deltaEnergy   = finalEnergy - stage0Energy;
            deltaRoute    = finalRoute - stage0Route;

            totalNewRoutes = nansum(newRouteVals);

            if abs(stage0Route) > tieTol
                sortieGrowthRatio = 100 * deltaRoute / stage0Route;
            else
                sortieGrowthRatio = NaN;
            end

            if abs(stage0Energy) > tieTol
                energyGrowthRatio = 100 * deltaEnergy / stage0Energy;
            else
                energyGrowthRatio = NaN;
            end

            if abs(finalRoute) > tieTol
                finalEnergyPerSortie = finalEnergy / finalRoute;
            else
                finalEnergyPerSortie = NaN;
            end

            rowsFinal(end+1, :) = {setNames{s}, instName, alg(a).name, ...
                stage0Sum, finalSum, stage0Energy, finalEnergy, stage0Route, finalRoute, ...
                totalNewRoutes, deltaSum, deltaEnergy, deltaRoute, ...
                sortieGrowthRatio, energyGrowthRatio, finalEnergyPerSortie, statusFlag}; %#ok<SAGROW>

            for k = 1:numel(stageVals)
                rowsStage(end+1, :) = {setNames{s}, instName, alg(a).name, ...
                    stageVals(k), sumVals(k), energyVals(k), routeVals(k), newRouteVals(k), statusFlag}; %#ok<SAGROW>
            end
        end
    end
end

finalTbl = cell2table(rowsFinal, 'VariableNames', ...
    {'Scale','Instance','Algorithm', ...
     'Stage0_SumObj','Final_SumObj','Stage0_DroneEnergy','Final_DroneEnergy', ...
     'Stage0_NumSorties','Final_NumSorties','Total_NewRoutes', ...
     'Delta_SumObj','Delta_DroneEnergy','Delta_NumSorties', ...
     'SortieGrowthRatio_pct','EnergyGrowthRatio_pct','FinalEnergyPerSortie', ...
     'StatusFlag'});

stageTbl = cell2table(rowsStage, 'VariableNames', ...
    {'Scale','Instance','Algorithm','Stage','SumObj','DroneEnergy','NumSorties','NewRouteCreated','StatusFlag'});

writetable(finalTbl, outXlsx, 'Sheet', 'Raw_FinalMetrics');
writetable(stageTbl, outXlsx, 'Sheet', 'Raw_StagewiseMetrics');

%% ========================= Table 1: Final-stage comparison ============
% Panel A: descriptive statistics of final objective
panelA_rows = {};
for g = 1:numel(groupOrder)
    grp = groupOrder{g};
    if strcmp(grp, 'Overall')
        Tg = finalTbl(strcmp(finalTbl.StatusFlag, "OK"), :);
    else
        Tg = finalTbl(strcmp(finalTbl.Scale, grp) & strcmp(finalTbl.StatusFlag, "OK"), :);
    end

    for a = 1:numel(algOrder)
        Ta = Tg(strcmp(Tg.Algorithm, algOrder{a}), :);
        if isempty(Ta), continue; end

        x = Ta.Final_SumObj;
        panelA_rows(end+1, :) = {grp, algOrder{a}, ...
            mean(x,'omitnan'), median(x,'omitnan'), std(x,'omitnan'), ...
            min(x), max(x), height(Ta)}; %#ok<SAGROW>
    end
end
Table1_PanelA = cell2table(panelA_rows, 'VariableNames', ...
    {'Scale','Algorithm','Mean_Final_SumObj','Median_Final_SumObj','Std_Final_SumObj','Min_Final_SumObj','Max_Final_SumObj','N'});

% Panel B: pairwise comparison against RH-BPC
panelB_rows = {};
for g = 1:numel(groupOrder)
    grp = groupOrder{g};
    if strcmp(grp, 'Overall')
        Tg = finalTbl(strcmp(finalTbl.StatusFlag, "OK"), :);
    else
        Tg = finalTbl(strcmp(finalTbl.Scale, grp) & strcmp(finalTbl.StatusFlag, "OK"), :);
    end

    Tref = Tg(strcmp(Tg.Algorithm, 'RH-BPC'), {'Instance','Final_SumObj'});
    Tref.Properties.VariableNames = {'Instance','RH_BPC_Final_SumObj'};

    baselines = {'Greedy','MIP','Repair'};
    for b = 1:numel(baselines)
        Tb = Tg(strcmp(Tg.Algorithm, baselines{b}), {'Instance','Final_SumObj'});
        Tb.Properties.VariableNames = {'Instance','Baseline_Final_SumObj'};

        Tpair = innerjoin(Tref, Tb, 'Keys', 'Instance');
        if isempty(Tpair)
            continue;
        end

        diffv = Tpair.RH_BPC_Final_SumObj - Tpair.Baseline_Final_SumObj;
        wins   = sum(diffv < -tieTol);
        ties   = sum(abs(diffv) <= tieTol);
        losses = sum(diffv > tieTol);
        nPair  = height(Tpair);
        winRate = 100 * wins / nPair;

        panelB_rows(end+1, :) = {grp, baselines{b}, wins, ties, losses, nPair, winRate}; %#ok<SAGROW>
    end
end
Table1_PanelB = cell2table(panelB_rows, 'VariableNames', ...
    {'Scale','Baseline','RH_BPC_Wins','Ties','Losses','N_Pairs','RH_BPC_WinRate_pct'});

% Panel C: best-solution frequency
panelC_rows = {};
for g = 1:numel(groupOrder)
    grp = groupOrder{g};
    if strcmp(grp, 'Overall')
        Tg = finalTbl(strcmp(finalTbl.StatusFlag, "OK"), :);
    else
        Tg = finalTbl(strcmp(finalTbl.Scale, grp) & strcmp(finalTbl.StatusFlag, "OK"), :);
    end

    instList = unique(Tg.Instance, 'stable');
    bestCount = zeros(1, numel(algOrder));

    for i = 1:numel(instList)
        Ti = Tg(strcmp(Tg.Instance, instList{i}), {'Algorithm','Final_SumObj'});
        if isempty(Ti), continue; end
        bestVal = min(Ti.Final_SumObj);
        isBest = abs(Ti.Final_SumObj - bestVal) <= tieTol;
        bestAlgs = Ti.Algorithm(isBest);
        for a = 1:numel(algOrder)
            bestCount(a) = bestCount(a) + sum(strcmp(bestAlgs, algOrder{a}));
        end
    end

    for a = 1:numel(algOrder)
        panelC_rows(end+1, :) = {grp, algOrder{a}, bestCount(a)}; %#ok<SAGROW>
    end
end
Table1_PanelC = cell2table(panelC_rows, 'VariableNames', ...
    {'Scale','Algorithm','Best_FinalObjective_Frequency'});

writetable(Table1_PanelA, outXlsx, 'Sheet', 'Table1_PanelA');
writetable(Table1_PanelB, outXlsx, 'Sheet', 'Table1_PanelB');
writetable(Table1_PanelC, outXlsx, 'Sheet', 'Table1_PanelC');

%% ========================= Table 2: Dynamic responsiveness ============
% Core panel
panel2A_rows = {};
for g = 1:numel(groupOrder)
    grp = groupOrder{g};
    if strcmp(grp, 'Overall')
        Tg = finalTbl(strcmp(finalTbl.StatusFlag, "OK"), :);
    else
        Tg = finalTbl(strcmp(finalTbl.Scale, grp) & strcmp(finalTbl.StatusFlag, "OK"), :);
    end

    for a = 1:numel(algOrder)
        Ta = Tg(strcmp(Tg.Algorithm, algOrder{a}), :);
        if isempty(Ta), continue; end

        panel2A_rows(end+1, :) = {grp, algOrder{a}, ...
            mean(Ta.Delta_SumObj,'omitnan'), median(Ta.Delta_SumObj,'omitnan'), std(Ta.Delta_SumObj,'omitnan'), ...
            mean(Ta.Delta_NumSorties,'omitnan'), median(Ta.Delta_NumSorties,'omitnan'), std(Ta.Delta_NumSorties,'omitnan'), ...
            mean(Ta.Delta_DroneEnergy,'omitnan'), median(Ta.Delta_DroneEnergy,'omitnan'), std(Ta.Delta_DroneEnergy,'omitnan'), ...
            mean(Ta.Total_NewRoutes,'omitnan'), median(Ta.Total_NewRoutes,'omitnan'), ...
            height(Ta)}; %#ok<SAGROW>
    end
end
Table2_PanelA = cell2table(panel2A_rows, 'VariableNames', ...
    {'Scale','Algorithm', ...
     'Mean_Delta_SumObj','Median_Delta_SumObj','Std_Delta_SumObj', ...
     'Mean_Delta_NumSorties','Median_Delta_NumSorties','Std_Delta_NumSorties', ...
     'Mean_Delta_DroneEnergy','Median_Delta_DroneEnergy','Std_Delta_DroneEnergy', ...
     'Mean_TotalNewRoutes','Median_TotalNewRoutes','N'});

% Auxiliary panel
panel2B_rows = {};
for g = 1:numel(groupOrder)
    grp = groupOrder{g};
    if strcmp(grp, 'Overall')
        Tg = finalTbl(strcmp(finalTbl.StatusFlag, "OK"), :);
    else
        Tg = finalTbl(strcmp(finalTbl.Scale, grp) & strcmp(finalTbl.StatusFlag, "OK"), :);
    end

    for a = 1:numel(algOrder)
        Ta = Tg(strcmp(Tg.Algorithm, algOrder{a}), :);
        if isempty(Ta), continue; end

        panel2B_rows(end+1, :) = {grp, algOrder{a}, ...
            mean(Ta.SortieGrowthRatio_pct,'omitnan'), ...
            mean(Ta.EnergyGrowthRatio_pct,'omitnan'), ...
            mean(Ta.FinalEnergyPerSortie,'omitnan'), ...
            median(Ta.FinalEnergyPerSortie,'omitnan'), ...
            height(Ta)}; %#ok<SAGROW>
    end
end
Table2_PanelB = cell2table(panel2B_rows, 'VariableNames', ...
    {'Scale','Algorithm', ...
     'Mean_SortieGrowthRatio_pct','Mean_EnergyGrowthRatio_pct', ...
     'Mean_FinalEnergyPerSortie','Median_FinalEnergyPerSortie','N'});

writetable(Table2_PanelA, outXlsx, 'Sheet', 'Table2_PanelA');
writetable(Table2_PanelB, outXlsx, 'Sheet', 'Table2_PanelB');

%% ========================= Table notes sheet ==========================
notes = {
    'Sheet','Description';
    'Raw_FinalMetrics','Instance-level final and increment metrics extracted from raw files';
    'Raw_StagewiseMetrics','Optional stage-wise raw metrics for checking';
    'Table1_PanelA','Final-stage objective descriptive statistics';
    'Table1_PanelB','Pairwise instance-level comparison against RH-BPC';
    'Table1_PanelC','Best final objective frequency';
    'Table2_PanelA','Core dynamic responsiveness statistics';
    'Table2_PanelB','Auxiliary growth/burden indicators'
    };
notesTbl = cell2table(notes(2:end,:), 'VariableNames', notes(1,:));
writetable(notesTbl, outXlsx, 'Sheet', 'README');

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
