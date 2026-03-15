function merge_and_stats(inputFile, outputFile)
    % Get all sheet names
    [~, sheets] = xlsfinfo(inputFile);

    % Store all data
    allData = table();

    % Iterate through each sheet and merge data
    for i = 1:numel(sheets)
        data = readtable(inputFile, 'Sheet', sheets{i}, 'VariableNamingRule', 'preserve');

        % Ensure required column names exist
        if ~all(ismember({'Group','Distance_nm','Force2x'}, data.Properties.VariableNames))
            warning('Sheet %s does not contain required columns.', sheets{i});
            continue;
        end

        % Add source information (optional)
        data.SourceSheet = repmat(string(sheets{i}), height(data), 1);

        % Merge into the overall table
        allData = [allData; data];
    end

    % ---- Sort merged data by Group start value ----
    groupStart_All = parseGroupStart(allData.Group);
    allData.Group = string(allData.Group); % Convert to string to avoid categorical sorting issues
    allData = addvars(allData, groupStart_All, 'Before', 'Group', 'NewVariableNames', 'GroupStart');
    allData = sortrows(allData, {'GroupStart','Group'}); % Sort first by numeric start, then by text for stability
    allData.GroupStart = []; % Remove temporary column used for sorting

    % Write merged data into one sheet
    writetable(allData, outputFile, 'Sheet', 'MergedData');

    % ---- Statistics section (sorted by Group start value) ----
    groups = string(unique(allData.Group));               % Unique groups
    groupStart_Uniq = parseGroupStart(groups);            % Parse start values for unique groups
    [~, orderIdx] = sort(groupStart_Uniq);                % Sorting index
    groups = groups(orderIdx);

    stats = table();

    for i = 1:numel(groups)
        grp = groups(i);
        grpData = allData(allData.Group == grp, :);

        % Distance_nm statistics
        dist_median = median(grpData.Distance_nm);
        dist_mean   = mean(grpData.Distance_nm);
        dist_iqr    = iqr(grpData.Distance_nm);  % Interquartile range

        % Force2x statistics
        force_mean = mean(grpData.Force2x);
        force_std  = std(grpData.Force2x);
        force_sem  = force_std / sqrt(height(grpData)); % Standard error of the mean (SEM)

        % Sample size
        count = height(grpData);

        % Add to statistics table
        stats = [stats; table(grp, dist_median, dist_mean, dist_iqr, ...
                              force_mean, force_std, force_sem, count, ...
                              'VariableNames', {'Group','Dist_Median','Dist_Mean','Dist_IQR', ...
                              'Force_Mean','Force_STD','Force_SEM','Count'})];
    end

    % Write statistics results
    writetable(stats, outputFile, 'Sheet', 'Statistics');

    fprintf('Merging and statistics completed, results saved to %s\n', outputFile);
end

function starts = parseGroupStart(groupVals)
    % Parse Group (e.g., "0-50") into numeric start value (0), used for numeric sorting
    g = string(groupVals);
    % Remove possible whitespace
    g = strtrim(g);
    % Split by "-" and take the first part
    parts = split(g, "-");
    starts = str2double(parts(:,1));
    % If parsing fails (NaN), set to Inf to avoid interfering with sorting
    starts(isnan(starts)) = Inf;
end
