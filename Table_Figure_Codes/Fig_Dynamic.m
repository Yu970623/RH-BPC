%% Online responsiveness: mean curve + shaded band of stage-wise Sum_obj
% For each customer scale (15 / 30 / 50), draw one figure.
% X-axis: 15 instances
% Y-axis: stage-wise Sum_obj statistics of each algorithm on that instance
% Line   : mean Sum_obj across stages
% Band   : min-max range across stages
%
% If you prefer a more robust band, replace min/max with prctile(25/75).

clear; clc; close all;

%% ========================= User settings ==============================
rootDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment';
outDir  = fullfile(rootDir, 'Analysis_OnlineResponsiveness');
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

% ---- Algorithm folders and filename suffixes ----
alg(1).name   = 'RH-BPC';
alg(1).folder = fullfile(rootDir, 'RH-BCP', 'DynamicPickup_details');
alg(1).suffix = '_RH-BPC.xlsx';

alg(2).name   = 'RH-Greedy';
alg(2).folder = fullfile(rootDir, 'Greedy', 'DynamicPickup_details_myopic');
alg(2).suffix = '_Greedy.xlsx';

alg(3).name   = 'RH-MIP';
alg(3).folder = fullfile(rootDir, 'MIP', 'DynamicPickup_details');
alg(3).suffix = '_MIP.xlsx';

alg(4).name   = 'RH-Repair';
alg(4).folder = fullfile(rootDir, 'Repair', 'DynamicPickup_details_repair');
alg(4).suffix = '_Repair.xlsx';

algColors = [
    0.129, 0.400, 0.674;   % RH-BPC
    0.850, 0.325, 0.098;   % Greedy
    0.494, 0.184, 0.556;   % MIP
    0.466, 0.674, 0.188    % Repair
];

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

allSets   = {Set1, Set2, Set3};
setNames  = {'15-customer instances', '30-customer instances', '50-customer instances'};
setTags   = {'15customers', '30customers', '50customers'};

%% ========================= Extract stage-wise data =====================
% Store statistics:
% stat{s, a}.meanVec  : mean Sum_obj across stages for 15 instances
% stat{s, a}.lowVec   : lower bound across stages (min or Q1)
% stat{s, a}.highVec  : upper bound across stages (max or Q3)
% stat{s, a}.incVec   : increment = high - low
% stat{s, a}.status   : file read status

stat = cell(numel(allSets), numel(alg));
summaryRows = {};

for s = 1:numel(allSets)
    curSet = allSets{s};
    for a = 1:numel(alg)

        meanVec = nan(1, numel(curSet));
        lowVec  = nan(1, numel(curSet));
        highVec = nan(1, numel(curSet));
        incVec  = nan(1, numel(curSet));
        nStageVec = nan(1, numel(curSet));
        statusCell = strings(1, numel(curSet));

        for i = 1:numel(curSet)
            instName = curSet{i};

            xlsxFile = fullfile(alg(a).folder, ['Dynamic_' instName alg(a).suffix]);
            csvFile  = strrep(xlsxFile, '.xlsx', '.csv');

            T = table();
            loaded = false;

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
                statusCell(i) = "Missing/Empty";
                continue;
            end

            colStage  = find_col(T, {'Stage'});
            colStatus = find_col(T, {'Status'});
            colSumObj = find_col(T, {'Sum_obj','Sum obj','SumObj'});

            if isempty(colStage) || isempty(colStatus) || isempty(colSumObj)
                statusCell(i) = "MissingColumn";
                continue;
            end

            statusVec = string(T.(colStatus));
            validMask = ~ismissing(statusVec) & strlength(strtrim(statusVec)) > 0;
            Tvalid = T(validMask, :);

            if isempty(Tvalid)
                statusCell(i) = "NoValidRow";
                continue;
            end

            % Prefer OK rows
            okMask = strcmpi(strtrim(string(Tvalid.(colStatus))), 'OK');
            if any(okMask)
                Tuse = Tvalid(okMask, :);
                statusCell(i) = "OK";
            else
                Tuse = Tvalid;
                statusCell(i) = "NonOK";
            end

            % Sort by stage
            Tuse = sortrows(Tuse, colStage);

            y = double(Tuse.(colSumObj));
            y = y(~isnan(y));

            if isempty(y)
                statusCell(i) = "NoSumObj";
                continue;
            end

            % ---- band definition ----
            % Option A: min-max range (more direct for cumulative growth)
            lowVec(i)  = min(y);
            highVec(i) = max(y);

            % ---- if you want Q1-Q3 band instead, use:
            % lowVec(i)  = prctile(y, 25);
            % highVec(i) = prctile(y, 75);

            meanVec(i) = mean(y, 'omitnan');
            incVec(i)  = highVec(i) - lowVec(i);
            nStageVec(i) = numel(y);

            summaryRows(end+1, :) = {setNames{s}, instName, alg(a).name, ...
                lowVec(i), meanVec(i), highVec(i), incVec(i), nStageVec(i), statusCell(i)}; %#ok<SAGROW>
        end

        stat{s, a}.meanVec = meanVec;
        stat{s, a}.lowVec  = lowVec;
        stat{s, a}.highVec = highVec;
        stat{s, a}.incVec  = incVec;
        stat{s, a}.nStageVec = nStageVec;
        stat{s, a}.status  = statusCell;
    end
end

summaryTbl = cell2table(summaryRows, 'VariableNames', ...
    {'SetName','Instance','Algorithm','StageMin_SumObj','StageMean_SumObj', ...
     'StageMax_SumObj','StageIncrement_SumObj','NumStages','StatusFlag'});

writetable(summaryTbl, fullfile(outDir, 'OnlineResponsiveness_SumObj_Summary.xlsx'));
writetable(summaryTbl, fullfile(outDir, 'OnlineResponsiveness_SumObj_Summary.csv'));

%% ========================= Plot 3 separate figures =====================
for s = 1:numel(allSets)
    curSet = allSets{s};
    x = 1:numel(curSet);

    fig = figure('Color','w', 'Position',[120, 120, 1200, 500]);
    ax = axes(fig); hold(ax, 'on'); box(ax, 'on');

    % draw shaded bands first
    for a = 1:numel(alg)
        lowVec  = stat{s, a}.lowVec;
        highVec = stat{s, a}.highVec;
        meanVec = stat{s, a}.meanVec;

        valid = ~(isnan(lowVec) | isnan(highVec) | isnan(meanVec));
        if ~any(valid)
            continue;
        end

        xv = x(valid);
        lv = lowVec(valid);
        hv = highVec(valid);

        fill([xv, fliplr(xv)], [hv, fliplr(lv)], algColors(a,:), ...
            'FaceAlpha', 0.30, 'EdgeColor', 'none');
    end

    % draw mean curves
    for a = 1:numel(alg)
        meanVec = stat{s, a}.meanVec;
        valid = ~isnan(meanVec);
        if ~any(valid)
            continue;
        end

        xv = x(valid);
        yv = meanVec(valid);

        plot(ax, xv, yv, '-o', ...
            'Color', algColors(a,:), ...
            'LineWidth', 1.8, ...
            'MarkerSize', 5, ...
            'MarkerFaceColor', algColors(a,:));
    end

    ax.XLim = [1, numel(curSet)];
    ax.XTick = x;
    ax.XTickLabel = curSet;
    xtickangle(ax, 45);

    ax.FontName = 'Times New Roman';
    ax.FontSize = 11;
    ax.LineWidth = 1.0;

    ylabel(ax, 'Stage-wise objective value', 'Interpreter', 'tex');
    % title(ax, ['Online responsiveness under ' setNames{s}], ...
    %     'FontWeight', 'normal', 'FontName', 'Times New Roman');

    legend(ax, {alg.name}, 'Location', 'northeast', 'Box', 'off');
    grid(ax, 'on');

    saveas(fig, fullfile(outDir, ['OnlineResponsiveness_SumObj_' setTags{s} '.png']));
    savefig(fig, fullfile(outDir, ['OnlineResponsiveness_SumObj_' setTags{s} '.fig']));
end

disp('------------------------------------------------------------');
disp('Done.');
disp(['Summary table: ' fullfile(outDir, 'OnlineResponsiveness_SumObj_Summary.xlsx')]);
disp(['Figures saved in: ' outDir]);
disp('------------------------------------------------------------');

%% ========================= Local functions =============================
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