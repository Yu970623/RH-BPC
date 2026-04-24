%% Generate RH-BPC stage-time summary table for 15/30/50-customer instances
% This script updates the table:
%   "The initial and dynamic planning time statistics ..."
% using RH-BPC result files directly.
%
% It updates the 15- and 30-customer groups and adds the 50-customer group.
%
% Outputs:
%   1) RHBPC_TimeStats_15_30_50.xlsx
%   2) RHBPC_TimeStats_15_30_50.csv
%   3) RHBPC_TimeStats_15_30_50.tex
%
% Table logic:
%   - Customer: parsed from instance name
%   - Stage: maximum stage index in the file
%   - Initial decision-making time: solve time at Stage 0
%   - Average decision time during dynamic programming:
%         mean solve time over stages > 0
%   - Route number change:
%         Stage0 NumSorties -> Final-stage NumSorties
%         If unchanged, show a single number
%   - Energy consumption of drone (Wh):
%         Stage0 DroneEnergy -> Final-stage DroneEnergy

clear; clc;

%% ========================= User settings ===============================
rootDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment';
inDir   = fullfile(rootDir, 'RH-BCP', 'DynamicPickup_details');
outDir  = fullfile(rootDir, 'Analysis_RHBPC_TimeStats');
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

outXlsx = fullfile(outDir, 'RHBPC_TimeStats_15_30_50.xlsx');
outCsv  = fullfile(outDir, 'RHBPC_TimeStats_15_30_50.csv');
outTex  = fullfile(outDir, 'RHBPC_TimeStats_15_30_50.tex');

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

%% ========================= Extract statistics ==========================
rows = {};

for s = 1:numel(allSets)
    curSet = allSets{s};

    for i = 1:numel(curSet)
        instName = curSet{i};
        xlsxFile = fullfile(inDir, ['Dynamic_' instName '_RH-BPC.xlsx']);
        csvFile  = fullfile(inDir, ['Dynamic_' instName '_RH-BPC.csv']);

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
            rows(end+1,:) = {instName, parse_customer_count(instName), NaN, '', '', '', '', "MissingFile"}; %#ok<SAGROW>
            continue;
        end

        colStage  = find_col(T, {'Stage'});
        colTime   = find_col(T, {'ReplanSolveTime(s)','ReplanSolveTime','SolveTime','Solve Ti','Time'});
        colRoute  = find_col(T, {'NumSorties','Num Sorties','NumSortie'});
        colEnergy = find_col(T, {'DroneEnergy(Wh)','DroneEnergy','Energy'});
        colStatus = find_col(T, {'Status'});

        if isempty(colStage) || isempty(colTime) || isempty(colRoute) || isempty(colEnergy)
            rows(end+1,:) = {instName, parse_customer_count(instName), NaN, '', '', '', '', "MissingColumn"}; %#ok<SAGROW>
            continue;
        end

        if ~isempty(colStatus)
            statusVec = string(T.(colStatus));
            validMask = ~ismissing(statusVec) & strlength(strtrim(statusVec)) > 0;
            T = T(validMask, :);
            if isempty(T)
                rows(end+1,:) = {instName, parse_customer_count(instName), NaN, '', '', '', '', "NoValidRow"}; %#ok<SAGROW>
                continue;
            end
            okMask = strcmpi(strtrim(string(T.(colStatus))), 'OK');
            if any(okMask)
                T = T(okMask, :);
            end
        end

        T = sortrows(T, colStage);

        stageVals  = double(T.(colStage));
        timeVals   = double(T.(colTime));
        routeVals  = double(T.(colRoute));
        energyVals = double(T.(colEnergy));

        if isempty(stageVals)
            rows(end+1,:) = {instName, parse_customer_count(instName), NaN, '', '', '', '', "NoStageData"}; %#ok<SAGROW>
            continue;
        end

        idx0 = find(stageVals == 0, 1, 'first');
        if isempty(idx0)
            idx0 = 1;
        end
        [finalStage, idxF] = max(stageVals);

        initTime = timeVals(idx0);

        dynMask = stageVals > stageVals(idx0);
        if any(dynMask)
            avgDynTime = mean(timeVals(dynMask), 'omitnan');
        else
            avgDynTime = initTime;
        end

        initRoute  = routeVals(idx0);
        finalRoute = routeVals(idxF);

        initEnergy  = energyVals(idx0);
        finalEnergy = energyVals(idxF);

        routeStr  = arrow_or_single(initRoute, finalRoute, 0);
        energyStr = arrow_or_single(initEnergy, finalEnergy, 3);

        rows(end+1,:) = {instName, parse_customer_count(instName), finalStage, ...
                         sprintf_num(initTime, 3, true), ...
                         sprintf_num(avgDynTime, 3, true), ...
                         routeStr, energyStr, statusFlag}; %#ok<SAGROW>
    end
end

resultTbl = cell2table(rows, 'VariableNames', ...
    {'Instances','Customer','Stage','Initial_decision_making_time', ...
     'Average_decision_time_during_dynamic_programming', ...
     'Route_number_change','Energy_consumption_of_drone_Wh','StatusFlag'});

writetable(resultTbl, outXlsx, 'Sheet', 'TimeStats');
writetable(resultTbl, outCsv);

%% ========================= Export LaTeX table ==========================
fid = fopen(outTex, 'w', 'n', 'UTF-8');
if fid < 0
    error('Cannot open output tex file: %s', outTex);
end

fprintf(fid, '%% Auto-generated LaTeX table for RH-BPC time statistics\n');
fprintf(fid, '\\begin{table*}[pos=h]\n');
fprintf(fid, '\t\\centering\n');
fprintf(fid, '\t\\caption{The initial and dynamic planning time statistics of all 15, 30, and 50 customer instances}\n');
fprintf(fid, '\t\\tiny\n');
fprintf(fid, '\t\\setlength{\\tabcolsep}{3pt}\n');
fprintf(fid, '\t\\renewcommand{\\arraystretch}{1}\n');
fprintf(fid, '\t\\begin{tabular}{ccccccc}\n');
fprintf(fid, '\t\t\\toprule\n');
fprintf(fid, '\t\tInstances & \\multicolumn{1}{l}{Customer} & \\multicolumn{1}{p{3em}}{Stage} & \\multicolumn{1}{p{12em}}{Initial decision-making time} & \\multicolumn{1}{p{16em}}{Average decision time during dynamic programming} & \\multicolumn{1}{p{16em}}{Route number change} & \\multicolumn{1}{p{16em}}{Energy consumption of drone (Wh)} \\\\\n');
fprintf(fid, '\t\t\\midrule\n');

for i = 1:height(resultTbl)
    fprintf(fid, '\t\t%s & %s & %s & %s & %s & %s & %s \\\\\n', ...
        latex_escape(resultTbl.Instances{i}), ...
        latex_escape(num2str(resultTbl.Customer(i))), ...
        latex_escape(num2str(resultTbl.Stage(i))), ...
        latex_escape(resultTbl.Initial_decision_making_time{i}), ...
        latex_escape(resultTbl.Average_decision_time_during_dynamic_programming{i}), ...
        latex_escape(resultTbl.Route_number_change{i}), ...
        latex_escape(resultTbl.Energy_consumption_of_drone_Wh{i}));
end

fprintf(fid, '\t\t\\bottomrule\n');
fprintf(fid, '\t\\end{tabular}%%\n');
fprintf(fid, '\t\\label{AllIns}\n');
fprintf(fid, '\\end{table*}\n');
fprintf(fid, '\\FloatBarrier\n');

fclose(fid);

disp('------------------------------------------------------------');
disp('Done. Files generated:');
disp(outXlsx);
disp(outCsv);
disp(outTex);
disp('------------------------------------------------------------');

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

function n = parse_customer_count(instName)
    toks = regexp(instName, ',(\d+)$', 'tokens', 'once');
    if isempty(toks)
        n = NaN;
    else
        n = str2double(toks{1});
    end
end

function s = sprintf_num(x, nd, add_s)
    if isnan(x)
        s = '';
        return;
    end
    fmt = ['%0.', num2str(nd), 'f'];
    v = sprintf(fmt, x);
    v = regexprep(v, '(\.\d*?[1-9])0+$', '$1');
    v = regexprep(v, '\.0+$', '');
    if add_s
        s = [v, 's'];
    else
        s = v;
    end
end

function s = arrow_or_single(a, b, nd)
    if isnan(a) || isnan(b)
        s = '';
        return;
    end
    if abs(a - b) <= 1e-9
        s = num_to_str(a, nd);
    else
        s = [num_to_str(a, nd), '→', num_to_str(b, nd)];
    end
end

function s = num_to_str(x, nd)
    fmt = ['%0.', num2str(nd), 'f'];
    s = sprintf(fmt, x);
    s = regexprep(s, '(\.\d*?[1-9])0+$', '$1');
    s = regexprep(s, '\.0+$', '');
end

function s = latex_escape(strIn)
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
