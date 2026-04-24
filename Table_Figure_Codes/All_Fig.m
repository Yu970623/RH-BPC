%% 5.3.1 Overall solution quality at the final stage
% Read dynamic-stage result files of all algorithms,
% extract final-stage metrics for each instance,
% export summary tables, and draw violin plots.
%
% Metrics extracted:
%   1) Final Sum_obj
%   2) Final DroneEnergy(Wh)
%   3) Total NewSortieCreated (sum over all stages)
%   4) Final NumSorties (optional, also exported)
%
% Author: ChatGPT
% -------------------------------------------------------------------------

clear; clc; close all;

%% =========================  User settings  =============================
rootDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment';
outDir  = fullfile(rootDir, 'Analysis_5_3_1');
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
setNames = {'Set1-15', 'Set2-30', 'Set3-50'};

% ---- Plot settings ----
algColors = [
    0.129, 0.400, 0.674;   % RH-BPC
    0.850, 0.325, 0.098;   % Greedy
    0.494, 0.184, 0.556;   % MIP
    0.466, 0.674, 0.188    % Repair
];

rng(1);  % only for jitter reproducibility if needed

%% =========================  Read all files  ============================
rows = {};

for s = 1:numel(allSets)
    curSet = allSets{s};
    for i = 1:numel(curSet)
        instName = curSet{i};

        for a = 1:numel(alg)
            fileName = ['Dynamic_' instName alg(a).suffix];
            filePath = fullfile(alg(a).folder, fileName);

            if ~isfile(filePath)
                warning('File not found: %s', filePath);
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, "MissingFile"}; %#ok<SAGROW>
                continue;
            end

            try
                T = readtable(filePath, 'VariableNamingRule', 'preserve');
            catch
                warning('Failed to read: %s', filePath);
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, "ReadError"}; %#ok<SAGROW>
                continue;
            end

            if isempty(T)
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, "EmptyTable"}; %#ok<SAGROW>
                continue;
            end

            % ---- Resolve column names robustly ----
            colStage       = find_col(T, {'Stage'});
            colStatus      = find_col(T, {'Status'});
            colSumObj      = find_col(T, {'Sum_obj','Sum obj','SumObj'});
            colEnergy      = find_col(T, {'DroneEnergy(Wh)','DroneEnergy','Energy'});
            colNewRoute    = find_col(T, {'NewSortieCreated','NewRouteCreated','New Routes','NewRoutes'});
            colNumSorties  = find_col(T, {'NumSorties','Num Sorties','NumSortie'});

            if isempty(colStage) || isempty(colStatus) || isempty(colSumObj) || isempty(colEnergy) || isempty(colNewRoute)
                warning('Required columns missing in: %s', filePath);
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, "MissingColumn"}; %#ok<SAGROW>
                continue;
            end

            stageVec  = T.(colStage);
            statusVec = string(T.(colStatus));

            % ---- keep valid rows ----
            validMask = ~ismissing(statusVec) & strlength(strtrim(statusVec)) > 0;
            Tvalid = T(validMask, :);

            if isempty(Tvalid)
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, NaN, NaN, NaN, "NoValidRow"}; %#ok<SAGROW>
                continue;
            end

            % ---- prefer rows with Status == OK ----
            statusValid = string(Tvalid.(colStatus));
            okMask = strcmpi(strtrim(statusValid), 'OK');

            if any(okMask)
                Tok = Tvalid(okMask, :);
                stageOk = Tok.(colStage);
                [~, idxLast] = max(stageOk);
                Tlast = Tok(idxLast, :);
                overallStatus = "OK";
            else
                % if no OK row exists, still take the largest-stage row
                stageValid = Tvalid.(colStage);
                [~, idxLast] = max(stageValid);
                Tlast = Tvalid(idxLast, :);
                overallStatus = "NonOKFinal";
            end

            finalSumObj   = Tlast.(colSumObj);
            finalEnergy   = Tlast.(colEnergy);
            if ~isempty(colNumSorties)
                finalSorties = Tlast.(colNumSorties);
            else
                finalSorties = NaN;
            end

            % ---- total new routes = sum over all valid rows ----
            newRouteVec = Tvalid.(colNewRoute);
            totalNewRoutes = nansum(double(newRouteVec));

            rows(end+1, :) = {setNames{s}, instName, alg(a).name, ...
                              double(finalSumObj), double(finalEnergy), ...
                              double(totalNewRoutes), double(finalSorties), overallStatus}; %#ok<SAGROW>
        end
    end
end

resultTbl = cell2table(rows, 'VariableNames', ...
    {'SetName','Instance','Algorithm','Final_SumObj','Final_DroneEnergy_Wh', ...
     'Total_NewRoutes','Final_NumSorties','StatusFlag'});

%% =========================  Export long table  =========================
longFile = fullfile(outDir, 'FinalMetrics_Long.xlsx');
writetable(resultTbl, longFile, 'Sheet', 'LongFormat');

%% =========================  Build wide table  ==========================
% One row per instance; columns grouped by algorithm
instList = unique(resultTbl(:, {'SetName','Instance'}), 'rows', 'stable');
wideTbl = instList;

metricNames = {'Final_SumObj','Final_DroneEnergy_Wh','Total_NewRoutes','Final_NumSorties'};

for a = 1:numel(alg)
    Ta = resultTbl(strcmp(resultTbl.Algorithm, alg(a).name), :);
    Ta = sortrows(Ta, {'SetName','Instance'});

    for m = 1:numel(metricNames)
        vn = [alg(a).name '_' metricNames{m}];
        wideTbl.(matlab.lang.makeValidName(vn)) = Ta.(metricNames{m});
    end
end

wideFile = fullfile(outDir, 'FinalMetrics_Wide.xlsx');
writetable(wideTbl, wideFile, 'Sheet', 'WideFormat');

%% =========================  Summary statistics  ========================
sumRows = {};

groupList = [setNames, {'Overall'}];

for g = 1:numel(groupList)
    if strcmp(groupList{g}, 'Overall')
        Tgrp = resultTbl(strcmp(resultTbl.StatusFlag, "OK"), :);
    else
        Tgrp = resultTbl(strcmp(resultTbl.SetName, groupList{g}) & strcmp(resultTbl.StatusFlag, "OK"), :);
    end

    for a = 1:numel(alg)
        Ta = Tgrp(strcmp(Tgrp.Algorithm, alg(a).name), :);
        if isempty(Ta)
            continue;
        end

        sumRows(end+1, :) = {groupList{g}, alg(a).name, ...
            mean(Ta.Final_SumObj,'omitnan'), std(Ta.Final_SumObj,'omitnan'), median(Ta.Final_SumObj,'omitnan'), ...
            mean(Ta.Final_DroneEnergy_Wh,'omitnan'), std(Ta.Final_DroneEnergy_Wh,'omitnan'), median(Ta.Final_DroneEnergy_Wh,'omitnan'), ...
            mean(Ta.Total_NewRoutes,'omitnan'), std(Ta.Total_NewRoutes,'omitnan'), median(Ta.Total_NewRoutes,'omitnan'), ...
            mean(Ta.Final_NumSorties,'omitnan'), std(Ta.Final_NumSorties,'omitnan'), median(Ta.Final_NumSorties,'omitnan'), ...
            height(Ta)}; %#ok<SAGROW>
    end
end

summaryTbl = cell2table(sumRows, 'VariableNames', ...
    {'Group','Algorithm', ...
     'Mean_SumObj','Std_SumObj','Median_SumObj', ...
     'Mean_Energy','Std_Energy','Median_Energy', ...
     'Mean_NewRoutes','Std_NewRoutes','Median_NewRoutes', ...
     'Mean_FinalNumSorties','Std_FinalNumSorties','Median_FinalNumSorties', ...
     'N'});

summaryFile = fullfile(outDir, 'FinalMetrics_Summary.xlsx');
writetable(summaryTbl, summaryFile, 'Sheet', 'Summary');

%% =========================  Compact table for paper ====================
% Optional compact strings, suitable for quick paper drafting
paperRows = {};
for g = 1:numel(groupList)
    if strcmp(groupList{g}, 'Overall')
        Tgrp = resultTbl(strcmp(resultTbl.StatusFlag, "OK"), :);
    else
        Tgrp = resultTbl(strcmp(resultTbl.SetName, groupList{g}) & strcmp(resultTbl.StatusFlag, "OK"), :);
    end

    rowCell = cell(1, 1 + numel(alg));
    rowCell{1} = groupList{g};

    for a = 1:numel(alg)
        Ta = Tgrp(strcmp(Tgrp.Algorithm, alg(a).name), :);
        if isempty(Ta)
            rowCell{1+a} = "";
        else
            rowCell{1+a} = sprintf('%.3f / %.1f / %.2f', ...
                mean(Ta.Final_SumObj,'omitnan'), ...
                mean(Ta.Final_DroneEnergy_Wh,'omitnan'), ...
                mean(Ta.Total_NewRoutes,'omitnan'));
        end
    end
    paperRows(end+1, :) = rowCell; %#ok<SAGROW>
end

paperTbl = cell2table(paperRows, 'VariableNames', ...
    ['Group', cellfun(@(x) matlab.lang.makeValidName(x), {alg.name}, 'UniformOutput', false)]);
writetable(paperTbl, summaryFile, 'Sheet', 'CompactForPaper');

%% =========================  Violin plots  ==============================
Tok = resultTbl(strcmp(resultTbl.StatusFlag, "OK"), :);

fig = figure('Color','w', 'Position',[100, 100, 1450, 420]);

metricList = {'Final_SumObj', 'Final_DroneEnergy_Wh', 'Total_NewRoutes'};
yLabels    = {'Final Sum\_obj', 'Final Drone Energy (Wh)', 'Total New Routes'};
titlesTxt  = {'(a) Final total objective', '(b) Final drone energy', '(c) Total new routes'};

for m = 1:numel(metricList)
    ax = subplot(1, 3, m); hold(ax, 'on'); box(ax, 'on');

    for a = 1:numel(alg)
        dataVec = Tok{strcmp(Tok.Algorithm, alg(a).name), metricList{m}};
        dataVec = dataVec(~isnan(dataVec));

        if isempty(dataVec)
            continue;
        end

        draw_violin(ax, dataVec, a, algColors(a,:), 0.35);

        % median marker
        plot(ax, a, median(dataVec,'omitnan'), 'ko', 'MarkerFaceColor','k', 'MarkerSize',5);

        % optional scatter
        xjit = a + 0.04 * (rand(size(dataVec)) - 0.5);
        scatter(ax, xjit, dataVec, 12, 'k', 'filled', ...
            'MarkerFaceAlpha', 0.20, 'MarkerEdgeAlpha', 0.20);
    end

    ax.XLim = [0.5, numel(alg) + 0.5];
    ax.XTick = 1:numel(alg);
    ax.XTickLabel = {alg.name};
    ax.FontName = 'Times New Roman';
    ax.FontSize = 11;
    ylabel(ax, yLabels{m}, 'Interpreter','tex');
    title(ax, titlesTxt{m}, 'FontWeight','normal');
    grid(ax, 'on');
end

sgtitle('Overall solution quality at the final stage', ...
    'FontName', 'Times New Roman', 'FontSize', 13, 'FontWeight', 'bold');

saveas(fig, fullfile(outDir, 'Violin_FinalMetrics.png'));
savefig(fig, fullfile(outDir, 'Violin_FinalMetrics.fig'));

%% =========================  Optional set-wise violin ===================
fig2 = figure('Color','w', 'Position',[100, 100, 1500, 820]);

for s = 1:numel(setNames)
    Tset = Tok(strcmp(Tok.SetName, setNames{s}), :);

    for m = 1:numel(metricList)
        ax = subplot(numel(setNames), numel(metricList), (s-1)*numel(metricList)+m);
        hold(ax, 'on'); box(ax, 'on');

        for a = 1:numel(alg)
            dataVec = Tset{strcmp(Tset.Algorithm, alg(a).name), metricList{m}};
            dataVec = dataVec(~isnan(dataVec));
            if isempty(dataVec), continue; end

            draw_violin(ax, dataVec, a, algColors(a,:), 0.32);
            plot(ax, a, median(dataVec,'omitnan'), 'ko', 'MarkerFaceColor','k', 'MarkerSize',4);
        end

        ax.XLim = [0.5, numel(alg) + 0.5];
        ax.XTick = 1:numel(alg);
        ax.XTickLabel = {alg.name};
        ax.FontName = 'Times New Roman';
        ax.FontSize = 10;
        if s == 1
            title(ax, metricList{m}, 'Interpreter','none', 'FontWeight','normal');
        end
        if m == 1
            ylabel(ax, setNames{s}, 'FontWeight','bold');
        end
        grid(ax, 'on');
    end
end

saveas(fig2, fullfile(outDir, 'Violin_FinalMetrics_BySet.png'));
savefig(fig2, fullfile(outDir, 'Violin_FinalMetrics_BySet.fig'));

disp('------------------------------------------------------------');
disp('Done.');
disp(['Long table   : ' longFile]);
disp(['Wide table   : ' wideFile]);
disp(['Summary file : ' summaryFile]);
disp(['Figure 1     : ' fullfile(outDir, 'Violin_FinalMetrics.png')]);
disp(['Figure 2     : ' fullfile(outDir, 'Violin_FinalMetrics_BySet.png')]);
disp('------------------------------------------------------------');

%% =========================  Local functions  ===========================
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

    % loose contains matching as backup
    for i = 1:numel(candNorm)
        idx = find(contains(varsNorm, candNorm{i}), 1, 'first');
        if ~isempty(idx)
            colName = vars{idx};
            return;
        end
    end
end

function out = normalize_names(in)
% Lowercase and remove spaces, underscores, brackets, punctuation.
    if ischar(in), in = {in}; end
    out = cell(size(in));
    for k = 1:numel(in)
        s = lower(string(in{k}));
        s = regexprep(s, '[\s_\-\(\)\[\]\{\},./\\]', '');
        out{k} = char(s);
    end
end

function draw_violin(ax, dataVec, xpos, faceColor, halfWidth)
% Simple violin plot without external toolbox
    dataVec = dataVec(~isnan(dataVec));
    if numel(unique(dataVec)) == 1
        % degenerate case: draw a slim box
        y0 = unique(dataVec);
        patch(ax, [xpos-0.05 xpos+0.05 xpos+0.05 xpos-0.05], ...
                 [y0-1e-6 y0-1e-6 y0+1e-6 y0+1e-6], faceColor, ...
                 'FaceAlpha', 0.35, 'EdgeColor', faceColor, 'LineWidth', 1.0);
        return;
    end

    [f, yi] = ksdensity(dataVec, 'Function', 'pdf');
    f = f ./ max(f) * halfWidth;

    xLeft  = xpos - f;
    xRight = xpos + f;

    patch(ax, [xLeft, fliplr(xRight)], [yi, fliplr(yi)], faceColor, ...
        'FaceAlpha', 0.35, 'EdgeColor', faceColor, 'LineWidth', 1.2);

    % quartile line
    q1 = prctile(dataVec, 25);
    q2 = prctile(dataVec, 50);
    q3 = prctile(dataVec, 75);

    plot(ax, [xpos-0.10 xpos+0.10], [q2 q2], 'k-', 'LineWidth', 1.5);
    plot(ax, [xpos xpos], [q1 q3], 'k-', 'LineWidth', 1.0);
end