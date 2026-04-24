%% Final Sum_obj violin plots for 15 / 30 / 50 customer instances
clear; clc; close all;

rootDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment';
outDir  = fullfile(rootDir, 'Analysis_5_3_1_FinalSumObj');
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
setTitles = {'15-customer instances', '30-customer instances', '50-customer instances'};
setFileTags = {'15customers', '30customers', '50customers'};

algColors = [
    0.129, 0.400, 0.674;   % RH-BPC
    0.850, 0.325, 0.098;   % Greedy
    0.494, 0.184, 0.556;   % MIP
    0.466, 0.674, 0.188    % Repair
];

rng(1);

rows = {};

for s = 1:numel(allSets)
    curSet = allSets{s};
    for i = 1:numel(curSet)
        instName = curSet{i};

        for a = 1:numel(alg)
            fileName = ['Dynamic_' instName alg(a).suffix];
            filePath = fullfile(alg(a).folder, fileName);

            if ~isfile(filePath)
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, "MissingFile"}; %#ok<SAGROW>
                continue;
            end

            try
                T = readtable(filePath, 'VariableNamingRule', 'preserve');
            catch
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, "ReadError"}; %#ok<SAGROW>
                continue;
            end

            if isempty(T)
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, "EmptyTable"}; %#ok<SAGROW>
                continue;
            end

            colStage  = find_col(T, {'Stage'});
            colStatus = find_col(T, {'Status'});
            colSumObj = find_col(T, {'Sum_obj','Sum obj','SumObj'});

            if isempty(colStage) || isempty(colStatus) || isempty(colSumObj)
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, "MissingColumn"}; %#ok<SAGROW>
                continue;
            end

            statusVec = string(T.(colStatus));
            validMask = ~ismissing(statusVec) & strlength(strtrim(statusVec)) > 0;
            Tvalid = T(validMask, :);

            if isempty(Tvalid)
                rows(end+1, :) = {setNames{s}, instName, alg(a).name, NaN, "NoValidRow"}; %#ok<SAGROW>
                continue;
            end

            statusValid = string(Tvalid.(colStatus));
            okMask = strcmpi(strtrim(statusValid), 'OK');

            if any(okMask)
                Tok = Tvalid(okMask, :);
                stageOk = Tok.(colStage);
                [~, idxLast] = max(stageOk);
                Tlast = Tok(idxLast, :);
                stat = "OK";
            else
                stageValid = Tvalid.(colStage);
                [~, idxLast] = max(stageValid);
                Tlast = Tvalid(idxLast, :);
                stat = "NonOKFinal";
            end

            finalSumObj = double(Tlast.(colSumObj));
            rows(end+1, :) = {setNames{s}, instName, alg(a).name, finalSumObj, stat}; %#ok<SAGROW>
        end
    end
end

resultTbl = cell2table(rows, 'VariableNames', ...
    {'SetName','Instance','Algorithm','Final_SumObj','StatusFlag'});

Tok = resultTbl(strcmp(resultTbl.StatusFlag, "OK"), :);

for s = 1:numel(setNames)
    Tset = Tok(strcmp(Tok.SetName, setNames{s}), :);

    fig = figure('Color','w', 'Position',[180, 180, 560, 460]);
    ax = axes(fig); hold(ax, 'on'); box(ax, 'on');

    for a = 1:numel(alg)
        dataVec = Tset{strcmp(Tset.Algorithm, alg(a).name), 'Final_SumObj'};
        dataVec = dataVec(~isnan(dataVec));
        if isempty(dataVec), continue; end

        draw_violin(ax, dataVec, a, algColors(a,:), 0.32);
        plot(ax, a, median(dataVec, 'omitnan'), 'ko', ...
            'MarkerFaceColor', 'k', 'MarkerSize', 5);

        xjit = a + 0.05 * (rand(size(dataVec)) - 0.5);
        scatter(ax, xjit, dataVec, 14, 'k', 'filled', ...
            'MarkerFaceAlpha', 0.18, 'MarkerEdgeAlpha', 0.18);
    end

    ax.XLim = [0.5, numel(alg) + 0.5];
    ax.XTick = 1:numel(alg);
    ax.XTickLabel = {alg.name};
    ax.FontName = 'Times New Roman';
    ax.FontSize = 12;
    ax.LineWidth = 1.0;
    ylabel(ax, 'Final objective value', 'Interpreter','tex');
    % title(ax, ['Overall solution quality at the final stage (' setTitles{s} ')'], ...
        % 'FontWeight','normal', 'FontName', 'Times New Roman');
    grid(ax, 'on');

    saveas(fig, fullfile(outDir, ['Violin_FinalSumObj_' setFileTags{s} '.png']));
    savefig(fig, fullfile(outDir, ['Violin_FinalSumObj_' setFileTags{s} '.fig']));
end

disp('Done.');

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

function draw_violin(ax, dataVec, xpos, faceColor, halfWidth)
    dataVec = dataVec(~isnan(dataVec));
    if numel(unique(dataVec)) == 1
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

    q1 = prctile(dataVec, 25);
    q2 = prctile(dataVec, 50);
    q3 = prctile(dataVec, 75);

    plot(ax, [xpos-0.10 xpos+0.10], [q2 q2], 'k-', 'LineWidth', 1.5);
    plot(ax, [xpos xpos], [q1 q3], 'k-', 'LineWidth', 1.0);
end