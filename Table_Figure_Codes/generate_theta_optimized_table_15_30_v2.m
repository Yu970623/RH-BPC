%% Optimized theta sensitivity table for RH-BPC (final version)
% This script implements the recommended table structure:
%
% Instance | N | Algorithm | Mean Objective values (Initial stage)
%          | Mean Objective values (Final stage)
%          | Mean Objective values (Delta)
%          | Activated Satellites
%          | Mean Drone routes (Initial stage)
%          | Mean Drone routes (Final stage)
%          | Mean NewSortieCreated
%          | Improved / Same / Worse vs. theta=1
%
% Data source:
%   D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\RH-BCP\theta1..theta5\DynamicPickup_details
%
% Scope:
%   Only 15- and 30-customer instances.
%
% Key logic:
%   1) Do NOT discard mean values.
%   2) Add Mean NewSortieCreated.
%   3) Add Improved / Same / Worse vs. theta=1 based on final-stage objective.
%   4) Activated Satellites = number of numeric tokens in First-echelon string - 1
%      Example: "1 4 5" -> 3 numbers -> 2 activated satellites.
%
% Output:
%   - Theta_Sensitivity_Optimized_Table.xlsx
%   - Theta_Sensitivity_Optimized_Table.csv
%   - Raw_Instance_Data / Summary_Numeric / Pairwise_vs_theta1 sheets in xlsx

clear; clc;

%% ========================= User settings ===============================
rootDir = 'D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\RH-BCP';
outDir  = fullfile(rootDir, 'Theta_Sensitivity_Analysis');
if ~exist(outDir, 'dir')
    mkdir(outDir);
end

outXlsx = fullfile(outDir, 'Theta_Sensitivity_Optimized_Table.xlsx');
outCsv  = fullfile(outDir, 'Theta_Sensitivity_Optimized_Table.csv');

thetaList = 1:5;
tieTol = 1e-6;

Set15 = {'Ca1-2,3,15', 'Ca1-3,5,15', 'Ca1-6,4,15', 'Ca2-2,3,15', 'Ca2-3,5,15', 'Ca2-6,4,15', ...
         'Ca3-2,3,15', 'Ca3-3,5,15', 'Ca3-6,4,15', 'Ca4-2,3,15', 'Ca4-3,5,15', 'Ca4-6,4,15', ...
         'Ca5-2,3,15', 'Ca5-3,5,15', 'Ca5-6,4,15'};

Set30 = {'Ca1-2,3,30', 'Ca1-3,5,30', 'Ca1-6,4,30', 'Ca2-2,3,30', 'Ca2-3,5,30', 'Ca2-6,4,30', ...
         'Ca3-2,3,30', 'Ca3-3,5,30', 'Ca3-6,4,30', 'Ca4-2,3,30', 'Ca4-3,5,30', 'Ca4-6,4,30', ...
         'Ca5-2,3,30', 'Ca5-3,5,30', 'Ca5-6,4,30'};

groupNames = {'15 Customers', '30 Customers'};
groupNs    = [15, 15];
groupSets  = {Set15, Set30};

%% ========================= Extract instance-level data =================
rawRows = {};

for g = 1:numel(groupSets)
    curSet = groupSets{g};
    curGroup = groupNames{g};
    curN = groupNs(g);

    for tt = 1:numel(thetaList)
        theta = thetaList(tt);
        inDir = fullfile(rootDir, ['theta', num2str(theta)], 'DynamicPickup_details');

        for i = 1:numel(curSet)
            instName = curSet{i};
            xlsxFile = fullfile(inDir, ['Dynamic_' instName '_RH-BPC.xlsx']);
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
                rawRows(end+1,:) = {curGroup, curN, theta, instName, ... %#ok<SAGROW>
                    NaN, NaN, NaN, NaN, NaN, NaN, NaN, "MissingFile"};
                continue;
            end

            colStage  = find_col(T, {'Stage'});
            colObj    = find_col(T, {'Sum_obj','Sum obj','SumObj'});
            colSort   = find_col(T, {'NumSorties','Num Sorties','NumSortie'});
            colFirst  = find_col(T, {'First-echelon','First echelon','FirstEchelon'});
            colNew    = find_col(T, {'NewSortieCreated','New Sortie Created','NewSortie'});
            colStatus = find_col(T, {'Status'});

            if isempty(colStage) || isempty(colObj) || isempty(colSort) || isempty(colFirst)
                rawRows(end+1,:) = {curGroup, curN, theta, instName, ... %#ok<SAGROW>
                    NaN, NaN, NaN, NaN, NaN, NaN, NaN, "MissingColumn"};
                continue;
            end

            % Prefer Status='OK' rows if Status exists
            if ~isempty(colStatus)
                statusVec = string(T.(colStatus));
                validMask = ~ismissing(statusVec) & strlength(strtrim(statusVec)) > 0;
                T = T(validMask, :);

                if isempty(T)
                    rawRows(end+1,:) = {curGroup, curN, theta, instName, ... %#ok<SAGROW>
                        NaN, NaN, NaN, NaN, NaN, NaN, NaN, "NoValidRow"};
                    continue;
                end

                okMask = strcmpi(strtrim(string(T.(colStatus))), 'OK');
                if any(okMask)
                    T = T(okMask, :);
                else
                    statusFlag = "NonOKFinal";
                end
            end

            T = sortrows(T, colStage);

            stageVals = double(T.(colStage));
            objVals   = double(T.(colObj));
            sortVals  = double(T.(colSort));

            if isempty(stageVals) || isempty(objVals) || isempty(sortVals)
                rawRows(end+1,:) = {curGroup, curN, theta, instName, ... %#ok<SAGROW>
                    NaN, NaN, NaN, NaN, NaN, NaN, NaN, "NoNumericData"};
                continue;
            end

            idx0 = find(stageVals == 0, 1, 'first');
            if isempty(idx0)
                idx0 = 1;
            end
            [~, idxF] = max(stageVals);

            initObj   = objVals(idx0);
            finalObj  = objVals(idxF);
            deltaObj  = finalObj - initObj;
            initSort  = sortVals(idx0);
            finalSort = sortVals(idxF);

            firstFinal = to_char_or_string(T.(colFirst)(idxF));
            actSat     = activated_satellites_from_first_echelon(firstFinal);

            if ~isempty(colNew)
                newVals = double(T.(colNew));
                totalNew = sum(newVals, 'omitnan');
            else
                totalNew = max(finalSort - initSort, 0);
            end

            rawRows(end+1,:) = {curGroup, curN, theta, instName, ... %#ok<SAGROW>
                initObj, finalObj, deltaObj, actSat, initSort, finalSort, totalNew, statusFlag};
        end
    end
end

rawTbl = cell2table(rawRows, 'VariableNames', ...
    {'InstanceGroup','N','theta','Instance', ...
     'InitialObj','FinalObj','DeltaObj', ...
     'ActivatedSatellites','InitialDroneRoutes','FinalDroneRoutes', ...
     'NewSortieCreated','StatusFlag'});

writetable(rawTbl, outXlsx, 'Sheet', 'Raw_Instance_Data');

%% ========================= Summary + pairwise vs theta=1 ==============
summaryRows = {};
pairRows = {};

for g = 1:numel(groupNames)
    curGroup = groupNames{g};
    curN = groupNs(g);

    Tg = rawTbl(strcmp(rawTbl.InstanceGroup, curGroup) & strcmp(rawTbl.StatusFlag, "OK"), :);

    % Reference theta=1
    Tref = Tg(Tg.theta == 1, {'Instance','FinalObj'});
    Tref.Properties.VariableNames = {'Instance','FinalObj_theta1'};

    for tt = 1:numel(thetaList)
        theta = thetaList(tt);
        Ttheta = Tg(Tg.theta == theta, :);

        meanInit   = mean(Ttheta.InitialObj, 'omitnan');
        meanFinal  = mean(Ttheta.FinalObj, 'omitnan');
        meanDelta  = mean(Ttheta.DeltaObj, 'omitnan');
        meanSat    = mean(Ttheta.ActivatedSatellites, 'omitnan');
        meanInitDr = mean(Ttheta.InitialDroneRoutes, 'omitnan');
        meanFinalDr= mean(Ttheta.FinalDroneRoutes, 'omitnan');
        meanNew    = mean(Ttheta.NewSortieCreated, 'omitnan');

        if theta == 1
            cmpStr = '/';
            improved = NaN; same = NaN; worse = NaN;
        else
            Tcmp = Ttheta(:, {'Instance','FinalObj'});
            Tcmp.Properties.VariableNames = {'Instance','FinalObj_theta'};
            Tpair = innerjoin(Tcmp, Tref, 'Keys', 'Instance');

            if isempty(Tpair)
                improved = NaN; same = NaN; worse = NaN;
                cmpStr = '--';
            else
                diffv = Tpair.FinalObj_theta - Tpair.FinalObj_theta1;
                improved = sum(diffv < -tieTol);       % theta better than theta=1
                same     = sum(abs(diffv) <= tieTol);
                worse    = sum(diffv > tieTol);        % theta worse than theta=1
                cmpStr   = sprintf('%d/%d/%d', improved, same, worse);
            end
        end

        summaryRows(end+1,:) = {curGroup, curN, ['theta=', num2str(theta)], ... %#ok<SAGROW>
            meanInit, meanFinal, meanDelta, meanSat, meanInitDr, meanFinalDr, meanNew, cmpStr};

        pairRows(end+1,:) = {curGroup, curN, theta, improved, same, worse, cmpStr}; %#ok<SAGROW>
    end
end

summaryTbl = cell2table(summaryRows, 'VariableNames', ...
    {'Instance','N','Algorithm', ...
     'MeanObjectiveInitial','MeanObjectiveFinal','MeanObjectiveDelta', ...
     'ActivatedSatellites','MeanDroneRoutesInitial','MeanDroneRoutesFinal', ...
     'MeanNewSortieCreated','ImprovedSameWorse_vs_theta1'});

pairTbl = cell2table(pairRows, 'VariableNames', ...
    {'InstanceGroup','N','theta','Improved','Same','Worse','SummaryString'});

%% ========================= Structured output ===========================
header = {'Instance','N','Algorithm', ...
    sprintf('Mean Objective values\n(Initial stage)'), ...
    sprintf('Mean Objective values\n(Final stage)'), ...
    sprintf('Mean Objective values\n(Delta)'), ...
    'Activated Satellites', ...
    sprintf('Mean Drone routes\n(Initial stage)'), ...
    sprintf('Mean Drone routes\n(Final stage)'), ...
    'Mean NewSortieCreated', ...
    sprintf('Improved / Same / Worse\nvs. theta=1')};

structured = header;

for g = 1:numel(groupNames)
    Ts = summaryTbl(strcmp(summaryTbl.Instance, groupNames{g}), :);

    for r = 1:height(Ts)
        if r == 1
            row = {Ts.Instance{r}, Ts.N(r), Ts.Algorithm{r}, ...
                numfmt(Ts.MeanObjectiveInitial(r),2), ...
                numfmt(Ts.MeanObjectiveFinal(r),2), ...
                numfmt(Ts.MeanObjectiveDelta(r),2), ...
                numfmt(Ts.ActivatedSatellites(r),2), ...
                numfmt(Ts.MeanDroneRoutesInitial(r),2), ...
                numfmt(Ts.MeanDroneRoutesFinal(r),2), ...
                numfmt(Ts.MeanNewSortieCreated(r),2), ...
                Ts.ImprovedSameWorse_vs_theta1{r}};
        else
            row = {'','', Ts.Algorithm{r}, ...
                numfmt(Ts.MeanObjectiveInitial(r),2), ...
                numfmt(Ts.MeanObjectiveFinal(r),2), ...
                numfmt(Ts.MeanObjectiveDelta(r),2), ...
                numfmt(Ts.ActivatedSatellites(r),2), ...
                numfmt(Ts.MeanDroneRoutesInitial(r),2), ...
                numfmt(Ts.MeanDroneRoutesFinal(r),2), ...
                numfmt(Ts.MeanNewSortieCreated(r),2), ...
                Ts.ImprovedSameWorse_vs_theta1{r}};
        end
        structured(end+1,:) = row; %#ok<SAGROW>
    end
end

writecell(structured, outXlsx, 'Sheet', 'Theta_Table');
writetable(summaryTbl, outXlsx, 'Sheet', 'Summary_Numeric');
writetable(pairTbl, outXlsx, 'Sheet', 'Pairwise_vs_theta1');

structuredTbl = cell2table(structured(2:end,:), 'VariableNames', matlab.lang.makeValidName(header));
writetable(structuredTbl, outCsv);

disp('============================================================');
disp('Done. Files generated:');
disp(outXlsx);
disp(outCsv);
disp('============================================================');

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

function s = to_char_or_string(x)
    if iscell(x)
        x = x{1};
    end
    if isstring(x)
        s = char(x);
    elseif ischar(x)
        s = x;
    else
        try
            s = char(string(x));
        catch
            s = '';
        end
    end
end

function n = activated_satellites_from_first_echelon(routeStr)
    if isempty(routeStr)
        n = NaN;
        return;
    end
    toks = regexp(routeStr, '\d+', 'match');
    if isempty(toks)
        n = NaN;
    else
        n = max(0, numel(toks) - 1);
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
