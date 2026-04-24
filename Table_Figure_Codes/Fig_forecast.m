%% ============================================================
% Plot line + shaded band by instance
% x-axis : 15 instances
% y-axis : objective value
% line   : mean objective over all stages for each instance
% shadow : min-max objective range from Stage 0 to Final stage
%
% Methods:
%   1) RH-BPC
%   2) RH-BPC-Forecast
%   3) RH-BPC-Perfect Forecast
%% ============================================================

clear; clc; close all;

%% ====================== User settings =======================
baseDir1 = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\RH-BCP\DynamicPickup_details';
baseDir2 = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\RH-BCP v2\DynamicPickup_details';
baseDir3 = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\RH-BCP v3\DynamicPickup_details';

outDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\Forecast_Comparison_15\Figures';
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

instances = {'Ca1-2,3,15', 'Ca1-3,5,15', 'Ca1-6,4,15', ...
             'Ca2-2,3,15', 'Ca2-3,5,15', 'Ca2-6,4,15', ...
             'Ca3-2,3,15', 'Ca3-3,5,15', 'Ca3-6,4,15', ...
             'Ca4-2,3,15', 'Ca4-3,5,15', 'Ca4-6,4,15', ...
             'Ca5-2,3,15', 'Ca5-3,5,15', 'Ca5-6,4,15'};

methods(1).name   = 'RH-BPC';
methods(1).folder = baseDir1;
methods(1).suffix = '_RH-BPC.xlsx';

methods(2).name   = 'RH-BPC-Forecast';
methods(2).folder = baseDir3;
methods(2).suffix = '.xlsx';   % 如果有特定后缀可自行改

methods(3).name   = 'RH-BPC-Perfect Forecast';
methods(3).folder = baseDir2;
methods(3).suffix = '.xlsx';   % 如果有特定后缀可自行改

% 颜色
methodColors = [
    0.0000, 0.4470, 0.7410;   % blue
    0.8500, 0.3250, 0.0980;   % orange
    0.4660, 0.6740, 0.1880    % green
];

%% ==================== Read instance-level stage stats ==================
nInst = numel(instances);
nMeth = numel(methods);

meanMat = nan(nMeth, nInst);
minMat  = nan(nMeth, nInst);
maxMat  = nan(nMeth, nInst);

for m = 1:nMeth
    for i = 1:nInst
        inst = instances{i};
        [objVec, statusFlag] = read_all_stage_objectives(methods(m).folder, inst, methods(m).suffix);

        if isempty(objVec)
            warning('No valid objective sequence found: %s | %s | %s', ...
                methods(m).name, inst, statusFlag);
            continue;
        end

        meanMat(m,i) = mean(objVec, 'omitnan');
        minMat(m,i)  = min(objVec);
        maxMat(m,i)  = max(objVec);
    end
end

%% ============================ Plot ====================================
x = 1:nInst;

fig = figure('Color', 'w', 'Position', [100 100 1200 520]);
hold on; box on;

for m = 1:nMeth
    yMean = meanMat(m,:);
    yMin  = minMat(m,:);
    yMax  = maxMat(m,:);

    valid = ~(isnan(yMean) | isnan(yMin) | isnan(yMax));
    xv = x(valid);
    yMeanv = yMean(valid);
    yMinv  = yMin(valid);
    yMaxv  = yMax(valid);

    if isempty(xv), continue; end

    % 阴影带：Stage 0 到 Final stage 的 min-max 区间
    fill([xv, fliplr(xv)], [yMinv, fliplr(yMaxv)], methodColors(m,:), ...
        'FaceAlpha', 0.16, 'EdgeColor', 'none');

    % 均值折线：所有 stage 的 objective 均值
    plot(xv, yMeanv, '-o', ...
        'Color', methodColors(m,:), ...
        'LineWidth', 2.0, ...
        'MarkerSize', 5, ...
        'DisplayName', methods(m).name);
end

xlim([1 nInst]);
xticks(x);
xticklabels(instances);
xtickangle(45);

xlabel('Instances', 'FontSize', 12);
ylabel('Objective value', 'FontSize', 12);
title('Stage-wise objective values across instances: mean curves with min-max bands', 'FontSize', 13);
legend('Location', 'northwest');
grid on;
set(gca, 'FontSize', 11);

% 导出
saveas(fig, fullfile(outDir, 'Forecast_LineShadow_ByInstance.png'));
saveas(fig, fullfile(outDir, 'Forecast_LineShadow_ByInstance.fig'));
exportgraphics(fig, fullfile(outDir, 'Forecast_LineShadow_ByInstance.pdf'), 'ContentType', 'vector');

disp('Figure exported to:');
disp(fullfile(outDir, 'Forecast_LineShadow_ByInstance.png'));
disp(fullfile(outDir, 'Forecast_LineShadow_ByInstance.pdf'));

%% ===================== Export plotting statistics ======================
header = [{'Instance'}, ...
    strcat(methods(1).name, '_Mean'), strcat(methods(1).name, '_Min'), strcat(methods(1).name, '_Max'), ...
    strcat(methods(2).name, '_Mean'), strcat(methods(2).name, '_Min'), strcat(methods(2).name, '_Max'), ...
    strcat(methods(3).name, '_Mean'), strcat(methods(3).name, '_Min'), strcat(methods(3).name, '_Max')];

outCell = cell(nInst+1, numel(header));
outCell(1,:) = header;

for i = 1:nInst
    outCell{i+1,1}  = instances{i};
    outCell{i+1,2}  = meanMat(1,i);
    outCell{i+1,3}  = minMat(1,i);
    outCell{i+1,4}  = maxMat(1,i);
    outCell{i+1,5}  = meanMat(2,i);
    outCell{i+1,6}  = minMat(2,i);
    outCell{i+1,7}  = maxMat(2,i);
    outCell{i+1,8}  = meanMat(3,i);
    outCell{i+1,9}  = minMat(3,i);
    outCell{i+1,10} = maxMat(3,i);
end

writecell(outCell, fullfile(outDir, 'Forecast_LineShadow_ByInstance_Stats.xlsx'));

%% ====================== Local functions ======================
function [objVec, statusFlag] = read_all_stage_objectives(folderPath, instName, suffix)

    objVec = [];
    statusFlag = "MissingFile";

    cand1 = fullfile(folderPath, ['Dynamic_' instName suffix]);
    cand2 = fullfile(folderPath, ['Dynamic_' instName '.xlsx']);
    cand3 = fullfile(folderPath, ['Dynamic_' instName '.csv']);

    filePath = '';
    if isfile(cand1)
        filePath = cand1;
    elseif isfile(cand2)
        filePath = cand2;
    elseif isfile(cand3)
        filePath = cand3;
    else
        D = dir(fullfile(folderPath, ['*' instName '*.xlsx']));
        if ~isempty(D)
            filePath = fullfile(folderPath, D(1).name);
        else
            D = dir(fullfile(folderPath, ['*' instName '*.csv']));
            if ~isempty(D)
                filePath = fullfile(folderPath, D(1).name);
            end
        end
    end

    if isempty(filePath)
        return;
    end

    T = table();
    try
        T = readtable(filePath, 'VariableNamingRule', 'preserve');
    catch
        statusFlag = "ReadError";
        return;
    end

    if isempty(T)
        statusFlag = "EmptyFile";
        return;
    end

    colStage  = find_col(T, {'Stage'});
    colObj    = find_col(T, {'Sum_obj','Sum obj','SumObj'});
    colStatus = find_col(T, {'Status'});

    if isempty(colStage) || isempty(colObj)
        statusFlag = "MissingColumn";
        return;
    end

    if ~isempty(colStatus)
        st = string(T.(colStatus));
        validMask = ~ismissing(st) & strlength(strtrim(st)) > 0;
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

    stageVec = double(T.(colStage));
    objVec   = double(T.(colObj));

    valid = ~(isnan(stageVec) | isnan(objVec));
    stageVec = stageVec(valid);
    objVec   = objVec(valid);

    if isempty(stageVec)
        statusFlag = "NoStageData";
        objVec = [];
        return;
    end

    [stageVec, idx] = sort(stageVec);
    objVec = objVec(idx);
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