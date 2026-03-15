function process_excel(inputFile, outputFile)
    % Read all sheet names from the Excel file
    [~, sheets] = xlsfinfo(inputFile);
    
    % Iterate through each sheet
    for i = 1:numel(sheets)
        % Read current sheet data, preserve original column names
        data = readtable(inputFile, 'Sheet', sheets{i}, 'VariableNamingRule', 'preserve');
        
        % Ensure required column names exist
        if ~all(ismember({'Distance 1','Force 2x'}, data.Properties.VariableNames))
            warning('Sheet %s does not contain required columns.', sheets{i});
            continue;
        end
        
        distance_um = data.("Distance 1");   % unit: μm
        force = data.("Force 2x");
        
        % Find the first sign change point (negative to positive)
        idx = find(force(1:end-1) < 0 & force(2:end) > 0, 1, 'first');
        
        % If no sign change, or all values are positive, skip
        if isempty(idx)
            fprintf('Sheet %s skipped (no sign change).\n', sheets{i});
            continue;
        end
        
        % Corresponding Distance value (μm)
        d0_um = distance_um(idx+1); 
        
        % Subtract d0 from all Distance values
        distance_shifted_um = distance_um - d0_um;
        
        % Keep only Distance >= 0
        validIdx = distance_shifted_um >= 0;
        distance_shifted_um = distance_shifted_um(validIdx);
        force_valid = force(validIdx);
        
        % Convert to nm
        distance_nm = distance_shifted_um * 1000;
        
        % Group into 50 nm intervals
        groupLabels = strcat(string(floor(distance_nm/50)*50), "-", ...
                             string(floor(distance_nm/50)*50 + 50));
        
        % Generate result table
        resultTable = table(groupLabels, distance_nm, force_valid, ...
            'VariableNames', {'Group','Distance_nm','Force2x'});
        
        % Write to output Excel file, keep sheet names
        writetable(resultTable, outputFile, 'Sheet', sheets{i});
        
        fprintf('Sheet %s processed.\n', sheets{i});
    end
    
    fprintf('All sheets processed. Output saved to %s\n', outputFile);
end
