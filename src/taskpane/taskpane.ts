/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import * as aq from 'arquero';
import vegaEmbed from 'vega-embed';
import mermaid from 'mermaid';

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Enable extended error logging
    OfficeExtension.config.extendedErrorLogging = true;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Arquero Event Listeners
    document.getElementById("open-arquero").onclick = () => tryCatch(openArqueroEditor);
    document.getElementById("close-arquero").onclick = closeArqueroEditor;
    document.getElementById("run-preview").onclick = runPreview;
    document.getElementById("save-output").onclick = outputToExcel;
    document.getElementById("table-select").onchange = (e) => {
        const tableName = (e.target as HTMLSelectElement).value;
        loadQueryForTable(tableName);
    };

    // Vega-Lite Event Listeners
    document.getElementById("open-vega").onclick = () => tryCatch(openVegaEditor);
    document.getElementById("close-vega").onclick = closeVegaEditor;
    document.getElementById("run-vega-preview").onclick = runVegaPreview;
    document.getElementById("save-vega-spec").onclick = saveVegaSpec;
    document.getElementById("save-vega-chart").onclick = saveVegaChart;
    document.getElementById("vega-table-select").onchange = onVegaTableSelect;
    document.getElementById("vega-spec-select").onchange = onVegaSpecSelect;

    // Mermaid Event Listeners
    document.getElementById("open-mermaid").onclick = () => tryCatch(openMermaidEditor);
    document.getElementById("close-mermaid").onclick = closeMermaidEditor;
    document.getElementById("run-mermaid-preview").onclick = runMermaidPreview;
    document.getElementById("save-mermaid-spec").onclick = saveMermaidSpec;
    document.getElementById("save-mermaid-chart").onclick = saveMermaidChart;
    document.getElementById("mermaid-spec-select").onchange = onMermaidSpecSelect;

    mermaid.initialize({ 
        startOnLoad: false,
        theme: 'base',
        themeVariables: {
            primaryColor: '#ffffff',
            primaryTextColor: '#000000',
            primaryBorderColor: '#000000',
            lineColor: '#000000',
            mainBkg: '#ffffff',
            nodeBorder: '#000000'
        },
        securityLevel: 'loose',
        fontFamily: 'Segoe UI, sans-serif',
        flowchart: { 
            htmlLabels: false,
            useMaxWidth: false
        },
        // Ensure htmlLabels is false globally for legacy graph syntax
        htmlLabels: false
    });
  }
});


/* Default helper for invoking an action and handling errors */
async function tryCatch(callback: ()=> Promise<void>):Promise<void> {
  try {
    const messageArea = document.getElementById("message-area");
    if (messageArea) messageArea.innerText = ""; // Clear previous messages
    await callback();
  } catch (error) {
    console.error(error);
    const messageArea = document.getElementById("message-area");
    if (messageArea) {
        messageArea.innerText = error instanceof Error ? error.message : String(error);
    }
  }
}

function showSpinner(message: string = "Processing...") {
    const overlay = document.getElementById("loading-overlay");
    const msg = document.getElementById("loading-message");
    if (overlay && msg) {
        msg.innerText = message;
        overlay.style.display = "flex";
    }
}

function hideSpinner() {
    const overlay = document.getElementById("loading-overlay");
    if (overlay) {
        overlay.style.display = "none";
    }
}

// -----------------------------------------------------------------------------
// Arquero Logic
// -----------------------------------------------------------------------------

function closeArqueroEditor() {
    document.getElementById("arquero-editor").style.display = "none";
}

async function openArqueroEditor(): Promise<void> {
    const editor = document.getElementById("arquero-editor");
    editor.style.display = "block";
    
    await loadTables();
    
    // Check if current selection is in a table and load its query
    await checkSelectionForQuery();

    document.getElementById("close-arquero").onclick = () => {
        editor.style.display = "none";
    };
    
    document.getElementById("run-preview").onclick = runPreview;
    document.getElementById("save-output").onclick = outputToExcel;
    
    // Add change listener to dropdown
    document.getElementById("table-select").onchange = (e) => {
        const tableName = (e.target as HTMLSelectElement).value;
        loadQueryForTable(tableName);
    };
}

async function checkSelectionForQuery() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            const tables = range.getTables(false); // Get tables overlapping with selection
            tables.load("items/name");
            await context.sync();

            if (tables.items.length > 0) {
                const tableName = tables.items[0].name;
                // Select in dropdown if it exists
                const select = document.getElementById("table-select") as HTMLSelectElement;
                if (select.querySelector(`option[value="${tableName}"]`)) {
                    select.value = tableName;
                    loadQueryForTable(tableName);
                }
            }
        });
    } catch (e) {
        console.error("Error checking selection:", e);
    }
}

async function loadQueryForTable(tableName: string) {
    const savedQuery = Office.context.document.settings.get(tableName);
    const textArea = document.getElementById("query-code") as HTMLTextAreaElement;
    
    if (savedQuery) {
        textArea.value = savedQuery;
    } else {
        // Optional: Clear textarea or keep previous? 
        // User might want to start fresh or keep editing. 
        // Let's clear it to indicate no query exists for this specific table, 
        // but only if we are switching context. 
        // Actually, safer to NOT clear if it's empty, so user doesn't lose work.
        // But if they select a table, they expect to see THAT table's query.
        // I'll clear it if null to avoid confusion.
        textArea.value = "";
    }
}

async function loadTables() {
    try {
        await Excel.run(async (context) => {
            const tables = context.workbook.tables;
            tables.load("items/name");
            await context.sync();

            const select = document.getElementById("table-select") as HTMLSelectElement;
            // Save current selection if any
            const currentVal = select.value;
            select.innerHTML = "";
            
            tables.items.forEach(table => {
                const option = document.createElement("option");
                option.value = table.name;
                option.text = table.name;
                select.appendChild(option);
            });
            
            if (currentVal && select.querySelector(`option[value="${currentVal}"]`)) {
                select.value = currentVal;
            }
        });
    } catch (error) {
        console.error(error);
    }
}

async function getArqueroTableFromExcel(tableName: string) {
    let data: any[] = [];
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(tableName);
        const range = table.getDataBodyRange();
        const headerRange = table.getHeaderRowRange();
        
        range.load("values");
        headerRange.load("values");
        await context.sync();

        const headers = headerRange.values[0];
        const rows = range.values;

        // Convert to array of objects for Arquero
        data = rows.map(row => {
            let obj = {};
            headers.forEach((header, index) => {
                obj[header] = row[index];
            });
            return obj;
        });
    });
    return aq.from(data);
}

async function runPreview() {
    showSpinner("Running query...");
    let code = (document.getElementById("query-code") as HTMLTextAreaElement).value;
    const previewArea = document.getElementById("preview-area");

    // Default to showing the full table if code is empty
    if (!code || code.trim() === "") {
        code = "dt";
    }

    try {
        const dt = await getTableData();
        
        // Execute user code
        const userFunc = new Function('dt', 'aq', `return ${code};`);
        const resultTable = userFunc(dt, aq);

        if (!resultTable) {
             throw new Error("Query did not return a table. Make sure your code returns an Arquero table.");
        }

        // Render HTML
        previewArea.innerHTML = resultTable.toHTML();
    } catch (error) {
        previewArea.innerText = "Error: " + error.message;
    } finally {
        hideSpinner();
    }
}

async function outputToExcel() {
    showSpinner("Exporting to Excel...");
    try {
    let code = (document.getElementById("query-code") as HTMLTextAreaElement).value;
    const outputTableName = (document.getElementById("output-table-name") as HTMLInputElement).value;
    const messageArea = document.getElementById("arquero-message");
    if (messageArea) messageArea.innerText = "";

    // Default to showing the full table if code is empty
    if (!code || code.trim() === "") {
        code = "dt";
    }

    if (!outputTableName) {
        if (messageArea) messageArea.innerText = "Please enter a unique name for the output table.";
        return;
    }

    // Validate table name (Excel rules: Start with letter/underscore, no spaces)
    if (!/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(outputTableName)) {
        if (messageArea) messageArea.innerText = "Error: Table name must start with a letter or underscore and contain only letters, numbers, and underscores (no spaces).";
        return;
    }

    try {
        const dt = await getTableData();
        const userFunc = new Function('dt', 'aq', `return ${code};`);
        const resultTable = userFunc(dt, aq);

        if (!resultTable) {
             throw new Error("Query did not return a table. Make sure your code returns an Arquero table.");
        }

        // Check row count limit
        if (resultTable.numRows() > 1000000) {
            if (messageArea) messageArea.innerText = `Error: Result has ${resultTable.numRows()} rows. Excel limit is 1,048,576 rows. Please refine your query.`;
            return;
        }

        // Convert back to array of arrays
        const objects = resultTable.objects();
        if (objects.length === 0) {
            if (messageArea) messageArea.innerText = "Query returned no data.";
            return;
        }

        const headers = Object.keys(objects[0]);
        // Sanitize values for Excel (convert objects/arrays to strings, handle undefined/null/NaN)
        const values = objects.map(obj => headers.map(h => {
            let val = obj[h];
            
            // Handle undefined or null
            if (val === undefined || val === null) {
                return ""; 
            }
            
            // Handle NaN
            if (typeof val === 'number' && isNaN(val)) {
                return "";
            }

            // Handle Objects/Arrays (excluding Date)
            if (typeof val === 'object' && !(val instanceof Date)) {
                try {
                    val = JSON.stringify(val);
                } catch (e) {
                    val = "[Circular/Error]";
                }
            }
            
            // Handle Strings: Length Limit & Control Characters
            if (typeof val === 'string') {
                // Remove invalid control characters (0x00-0x08, 0x0B-0x0C, 0x0E-0x1F)
                // Keep: Tab (0x09), New Line (0x0A), Carriage Return (0x0D)
                val = val.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");

                if (val.length > 32000) {
                    return val.substring(0, 32000) + "...(TRUNCATED)";
                }
            }
            
            return val;
        }));

        await Excel.run(async (context) => {
            // Check if table name already exists
            try {
                const existingTable = context.workbook.tables.getItem(outputTableName);
                existingTable.load("name");
                await context.sync();
                // If we get here, table exists. 
                // We could overwrite or error. User asked for "unique name".
                // Let's error for safety or maybe overwrite? 
                // "The output should be made a table in Excel with the unique name that the user provided."
                // Usually implies creating a new one.
                // I'll throw an error if it exists to be safe.
                throw new Error(`Table '${outputTableName}' already exists. Please choose a different name.`);
            } catch (e) {
                // ItemNotFound is good, means we can create it.
                if (e instanceof OfficeExtension.Error && e.code === "ItemNotFound") {
                    // Proceed
                } else if (e.message.indexOf("already exists") > -1) {
                     throw e;
                }
                 // Ignore other errors for now or rethrow?
            }

            const sheet = context.workbook.worksheets.add();
            
            // 1. Create Table with Headers
            const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
            headerRange.values = [headers];
            const table = sheet.tables.add(headerRange, true);
            table.name = outputTableName;
            
            // 2. Resize to full size and write data
            if (values.length > 0) {
                const fullRange = sheet.getRangeByIndexes(0, 0, values.length + 1, headers.length);
                table.resize(fullRange);
                
                // Write data in batches to avoid payload limits
                const batchSize = 500; // Reduced batch size further for safety
                for (let i = 0; i < values.length; i += batchSize) {
                    const batchValues = values.slice(i, i + batchSize);
                    // Use getRangeByIndexes on the SHEET, not relative to table, to be safe
                    const batchRange = sheet.getRangeByIndexes(1 + i, 0, batchValues.length, headers.length);
                    batchRange.values = batchValues;
                    
                    try {
                        // Sync periodically to flush data to Excel
                        await context.sync();
                    } catch (batchError) {
                        console.error(`Batch starting at index ${i} failed. Retrying row by row...`, batchError);
                        // Retry row by row for this batch
                        for (let j = 0; j < batchValues.length; j++) {
                            try {
                                const singleRowValues = [batchValues[j]];
                                const singleRowRange = sheet.getRangeByIndexes(1 + i + j, 0, 1, headers.length);
                                singleRowRange.values = singleRowValues;
                                await context.sync();
                            } catch (rowError) {
                                console.error(`Row ${i + j} failed:`, rowError);
                                console.error("Failed row data:", JSON.stringify(batchValues[j]));
                                // Try to write a placeholder to indicate failure
                                try {
                                    const errorRowRange = sheet.getRangeByIndexes(1 + i + j, 0, 1, 1); // First cell only
                                    errorRowRange.values = [["ERROR WRITING ROW"]];
                                    await context.sync();
                                } catch (e) {
                                    // If even that fails, just continue
                                }
                            }
                        }
                    }
                }
            }
            
            sheet.activate();
            await context.sync();
        });

        // Save the query associated with this table name (Outside Excel.run)
        try {
            await new Promise<void>((resolve, reject) => {
                Office.context.document.settings.set(outputTableName, code);
                Office.context.document.settings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.error("Failed to save settings:", result.error);
                        if (messageArea) messageArea.innerText += " (Warning: Failed to save query)";
                        resolve(); 
                    } else {
                        console.log("Settings saved successfully");
                        resolve();
                    }
                });
            });
        } catch (settingsError) {
             console.error("Settings Error:", settingsError);
        }
        
        if (messageArea) messageArea.innerText = `Success! Table '${outputTableName}' created.`;
        
        // Refresh table list so the new table appears in dropdown
        await loadTables();

    } catch (error) {
        console.error("Main Error:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug Info:", JSON.stringify(error.debugInfo));
        }
        if (messageArea) messageArea.innerText = "Error: " + error.message;
    }
    } finally {
        hideSpinner();
    }
}

async function getTableData() {
    const fileInput = document.getElementById("file-input") as HTMLInputElement;
    const tableName = (document.getElementById("table-select") as HTMLSelectElement).value;

    if (fileInput.files && fileInput.files.length > 0) {
        return await getArqueroTableFromFiles(fileInput.files);
    } else if (tableName) {
        return await getArqueroTableFromExcel(tableName);
    }
    throw new Error("Please select a table or upload files.");
}

async function getArqueroTableFromFiles(files: FileList) {
    let combinedData: any[] = [];
    
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const text = await readFileAsText(file);
        
        let table;
        if (file.name.toLowerCase().endsWith('.csv')) {
            table = aq.fromCSV(text);
        } else if (file.name.toLowerCase().endsWith('.json')) {
            const json = JSON.parse(text);
            table = aq.from(json);
        }
        
        if (table) {
            // Ensure data is flat objects, not nested
            const objects = table.objects();
            // Sanitize objects to ensure they are compatible with Excel (no nested objects/arrays)
            const sanitized = objects.map(obj => {
                const newObj = {};
                for (const key in obj) {
                    const val = obj[key];
                    if (typeof val === 'object' && val !== null && !(val instanceof Date)) {
                        newObj[key] = JSON.stringify(val);
                    } else {
                        newObj[key] = val;
                    }
                }
                return newObj;
            });
            combinedData = combinedData.concat(sanitized);
        }
    }
    
    return aq.from(combinedData);
}

function readFileAsText(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result as string);
        reader.onerror = (e) => reject(e);
        reader.readAsText(file);
    });
}

async function sortTable(tableName: string, columnName: string, direction: string) {
    await Excel.run(async (context) => {
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const table = currentWorksheet.tables.getItem(tableName);
        const columns = table.columns.load("items");
        await context.sync();

        // Find the index of the column
        let columnIndex = -1;
        for (let i = 0; i < columns.items.length; i++) {
            if (columns.items[i].name === columnName) {
                columnIndex = i;
                break;
            }
        }

        if (columnIndex !== -1) {
            const ascending = direction === "ascending";
            table.sort.apply([
                {
                    key: columnIndex,
                    ascending: ascending
                }
            ]);
        }

        await context.sync();
    });
}




// async function addFormulaToCell (formula:string):Promise<void> {
//   await Excel.run(async (context) => {
//     const range = context.workbook.getSelectedRange();
//     const cell = range.getCell(0,0);

//     cell.load(["values", "address"]);
//     await context.sync();

//     const cellValue = cell.values[0][0];

//     // Check if cell is not empty (Excel returns "" for empty cells)
//     if (cellValue !== "" && cellValue !== null && cellValue !== undefined) {
//       console.log(`Cell ${cell.address} has value: ${cellValue}`);
//       return cell;
//     }

//     console.log("Cell is empty.");
//     return null;

//   });
// }

// Vega-Lite Logic

async function openVegaEditor(): Promise<void> {
    const editor = document.getElementById("vega-editor");
    editor.style.display = "block";
    
    await loadVegaTables();
    
    // Check if current selection is in a table and load its specs
    await checkSelectionForVega();
}

function closeVegaEditor() {
    document.getElementById("vega-editor").style.display = "none";
}

async function loadVegaTables() {
    try {
        await Excel.run(async (context) => {
            const tables = context.workbook.tables;
            tables.load("items/name");
            await context.sync();

            const select = document.getElementById("vega-table-select") as HTMLSelectElement;
            const currentVal = select.value;
            select.innerHTML = "";
            
            // Add empty option
            const emptyOption = document.createElement("option");
            emptyOption.value = "";
            emptyOption.text = "-- Select Table --";
            select.appendChild(emptyOption);

            tables.items.forEach(table => {
                const option = document.createElement("option");
                option.value = table.name;
                option.text = table.name;
                select.appendChild(option);
            });
            
            if (currentVal && select.querySelector(`option[value="${currentVal}"]`)) {
                select.value = currentVal;
            }
        });
    } catch (error) {
        console.error(error);
    }
}

async function checkSelectionForVega() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            const tables = range.getTables(false);
            tables.load("items/name");
            await context.sync();

            if (tables.items.length > 0) {
                const tableName = tables.items[0].name;
                const select = document.getElementById("vega-table-select") as HTMLSelectElement;
                if (select.querySelector(`option[value="${tableName}"]`)) {
                    select.value = tableName;
                    await onVegaTableSelect();
                }
            }
        });
    } catch (e) {
        console.error("Error checking selection:", e);
    }
}

async function onVegaTableSelect() {
    const tableName = (document.getElementById("vega-table-select") as HTMLSelectElement).value;
    if (!tableName) return;
    
    await loadVegaSpecs(tableName);
}

async function loadVegaSpecs(tableName: string) {
    const settingsKey = `VEGA_SPECS_${tableName}`;
    const savedSpecs = Office.context.document.settings.get(settingsKey);
    const select = document.getElementById("vega-spec-select") as HTMLSelectElement;
    select.innerHTML = "";
    
    const newOption = document.createElement("option");
    newOption.value = "";
    newOption.text = "-- New Chart --";
    select.appendChild(newOption);

    if (savedSpecs) {
        const specs = JSON.parse(savedSpecs); // Array of {id, name, spec}
        specs.forEach((s: any) => {
            const option = document.createElement("option");
            option.value = s.id;
            option.text = s.name;
            select.appendChild(option);
        });
    }
    
    // Reset editor for new chart
    (document.getElementById("vega-chart-name") as HTMLInputElement).value = "";
    (document.getElementById("vega-spec") as HTMLTextAreaElement).value = JSON.stringify({
        "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
        "mark": "bar",
        "encoding": {
            "x": {"field": "Category", "type": "nominal"},
            "y": {"field": "Amount", "type": "quantitative"}
        }
    }, null, 2);
}

async function onVegaSpecSelect() {
    const tableName = (document.getElementById("vega-table-select") as HTMLSelectElement).value;
    const specId = (document.getElementById("vega-spec-select") as HTMLSelectElement).value;
    
    if (!tableName) return;

    if (!specId) {
        // New Chart
        (document.getElementById("vega-chart-name") as HTMLInputElement).value = "";
        // Keep default or clear?
        return;
    }

    const settingsKey = `VEGA_SPECS_${tableName}`;
    const savedSpecs = Office.context.document.settings.get(settingsKey);
    if (savedSpecs) {
        const specs = JSON.parse(savedSpecs);
        const spec = specs.find((s: any) => s.id === specId);
        if (spec) {
            (document.getElementById("vega-chart-name") as HTMLInputElement).value = spec.name;
            (document.getElementById("vega-spec") as HTMLTextAreaElement).value = JSON.stringify(spec.spec, null, 2);
        }
    }
}

async function runVegaPreview() {
    showSpinner("Generating preview...");
    try {
    const tableName = (document.getElementById("vega-table-select") as HTMLSelectElement).value;
    const specStr = (document.getElementById("vega-spec") as HTMLTextAreaElement).value;
    const messageArea = document.getElementById("vega-message");
    if (messageArea) messageArea.innerText = "Generating preview...";

    if (!tableName) {
        if (messageArea) messageArea.innerText = "Please select a source table.";
        return;
    }

    try {
        console.log("Parsing spec...");
        const spec = JSON.parse(specStr);
        
        console.log(`Fetching data for table: ${tableName}`);
        // Get Data
        const dt = await getArqueroTableFromExcel(tableName);
        const dataObjects = dt.objects();
        console.log(`Data fetched. Rows: ${dataObjects.length}`);
        
        // Inject data if not present or if placeholder used
        if (!spec.data || spec.data.name === "table") {
             spec.data = { values: dataObjects };
        }

        console.log("Embedding chart...");
        // Render
        await vegaEmbed('#vega-preview-area', spec, { actions: false });
        console.log("Chart rendered.");
        if (messageArea) messageArea.innerText = "";

    } catch (e) {
        console.error("Vega Preview Error:", e);
        if (messageArea) messageArea.innerText = "Error rendering chart: " + e.message;
    }
    } finally {
        hideSpinner();
    }
}
// Mermaid Logic

async function openMermaidEditor(): Promise<void> {
    const editor = document.getElementById("mermaid-editor");
    editor.style.display = "block";
    await loadMermaidSpecs();
}

function closeMermaidEditor() {
    document.getElementById("mermaid-editor").style.display = "none";
}

async function loadMermaidSpecs() {
    const settingsKey = "MERMAID_SPECS";
    const savedSpecs = Office.context.document.settings.get(settingsKey);
    const select = document.getElementById("mermaid-spec-select") as HTMLSelectElement;
    select.innerHTML = "";
    
    const newOption = document.createElement("option");
    newOption.value = "";
    newOption.text = "-- New Diagram --";
    select.appendChild(newOption);

    if (savedSpecs) {
        const specs = JSON.parse(savedSpecs); // Array of {id, name, code}
        specs.forEach((s: any) => {
            const option = document.createElement("option");
            option.value = s.id;
            option.text = s.name;
            select.appendChild(option);
        });
    }
    
    // Reset editor
    (document.getElementById("mermaid-chart-name") as HTMLInputElement).value = "";
    (document.getElementById("mermaid-code") as HTMLTextAreaElement).value = "flowchart TD;\n    A-->B;\n    A-->C;\n    B-->D;\n    C-->D;";
}

async function onMermaidSpecSelect() {
    const specId = (document.getElementById("mermaid-spec-select") as HTMLSelectElement).value;
    
    if (!specId) {
        // New Diagram
        (document.getElementById("mermaid-chart-name") as HTMLInputElement).value = "";
        (document.getElementById("mermaid-code") as HTMLTextAreaElement).value = "flowchart TD;\n    A-->B;\n    A-->C;\n    B-->D;\n    C-->D;";
        return;
    }

    const settingsKey = "MERMAID_SPECS";
    const savedSpecs = Office.context.document.settings.get(settingsKey);
    if (savedSpecs) {
        const specs = JSON.parse(savedSpecs);
        const spec = specs.find((s: any) => s.id === specId);
        if (spec) {
            (document.getElementById("mermaid-chart-name") as HTMLInputElement).value = spec.name;
            (document.getElementById("mermaid-code") as HTMLTextAreaElement).value = spec.code;
        }
    }
}

async function runMermaidPreview() {
    showSpinner("Generating preview...");
    const code = (document.getElementById("mermaid-code") as HTMLTextAreaElement).value;
    const messageArea = document.getElementById("mermaid-message");
    const previewArea = document.getElementById("mermaid-preview-area");
    
    if (messageArea) messageArea.innerText = "";
    previewArea.innerHTML = "";

    try {
        const { svg } = await mermaid.render('mermaid-svg-' + Date.now(), code);
        previewArea.innerHTML = svg;
    } catch (e) {
        console.error("Mermaid Preview Error:", e);
        if (messageArea) messageArea.innerText = "Error rendering diagram: " + e.message;
    } finally {
        hideSpinner();
    }
}

async function saveMermaidSpec() {
    const chartName = (document.getElementById("mermaid-chart-name") as HTMLInputElement).value;
    const code = (document.getElementById("mermaid-code") as HTMLTextAreaElement).value;
    const messageArea = document.getElementById("mermaid-message");
    if (messageArea) messageArea.innerText = "";

    if (!chartName) {
        if (messageArea) messageArea.innerText = "Please enter a diagram name.";
        return;
    }

    try {
        const settingsKey = "MERMAID_SPECS";
        let specs = [];
        const savedSpecs = Office.context.document.settings.get(settingsKey);
        if (savedSpecs) {
            specs = JSON.parse(savedSpecs);
        }

        // Check if updating existing
        const specId = (document.getElementById("mermaid-spec-select") as HTMLSelectElement).value;
        
        if (specId) {
            const index = specs.findIndex((s: any) => s.id === specId);
            if (index > -1) {
                specs[index].name = chartName;
                specs[index].code = code;
            }
        } else {
            // Create new
            const newId = Date.now().toString();
            specs.push({ id: newId, name: chartName, code: code });
            
            // Update dropdown
            const select = document.getElementById("mermaid-spec-select") as HTMLSelectElement;
            const option = document.createElement("option");
            option.value = newId;
            option.text = chartName;
            select.appendChild(option);
            select.value = newId;
        }

        Office.context.document.settings.set(settingsKey, JSON.stringify(specs));
        Office.context.document.settings.saveAsync((result) => {
             if (result.status === Office.AsyncResultStatus.Failed) {
                 if (messageArea) messageArea.innerText = "Failed to save diagram.";
             } else {
                 if (messageArea) messageArea.innerText = "Diagram saved successfully.";
             }
        });

    } catch (e) {
        if (messageArea) messageArea.innerText = "Error saving diagram.";
    }
}

async function saveMermaidChart() {
    showSpinner("Inserting diagram...");
    const messageArea = document.getElementById("mermaid-message");
    if (messageArea) messageArea.innerText = "";

    // First save spec
    await saveMermaidSpec();
    
    try {
        const code = (document.getElementById("mermaid-code") as HTMLTextAreaElement).value;
        
        console.log("Rendering SVG...");
        // Create a temporary container in the visible DOM to ensure text measurement works
        const tempContainer = document.createElement('div');
        tempContainer.style.visibility = 'hidden'; // Hide it but keep it in layout
        document.getElementById('app-body').appendChild(tempContainer);
        
        let svg;
        try {
            const result = await mermaid.render('mermaid-svg-export-' + Date.now(), code, tempContainer);
            svg = result.svg;
        } finally {
            if (tempContainer.parentNode) {
                tempContainer.parentNode.removeChild(tempContainer);
            }
        }
        
        // Encode SVG to base64
        const base64 = btoa(unescape(encodeURIComponent(svg)));
        
        console.log("Inserting image into Excel...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const image = sheet.shapes.addImage(base64);
            image.name = (document.getElementById("mermaid-chart-name") as HTMLInputElement).value || "MermaidDiagram";
            await context.sync();
        });
        
        if (messageArea) messageArea.innerText = "Diagram inserted into worksheet.";
        
    } catch (e) {
        console.error("Save Chart Error:", e);
        if (messageArea) messageArea.innerText = "Error inserting diagram: " + e.message;
    } finally {
        hideSpinner();
    }
}
async function saveVegaSpec() {
    const tableName = (document.getElementById("vega-table-select") as HTMLSelectElement).value;
    const chartName = (document.getElementById("vega-chart-name") as HTMLInputElement).value;
    const specStr = (document.getElementById("vega-spec") as HTMLTextAreaElement).value;
    const messageArea = document.getElementById("vega-message");
    if (messageArea) messageArea.innerText = "";

    if (!tableName || !chartName) {
        if (messageArea) messageArea.innerText = "Please select a table and enter a chart name.";
        return;
    }

    try {
        const spec = JSON.parse(specStr);
        const settingsKey = `VEGA_SPECS_${tableName}`;
        let specs = [];
        const savedSpecs = Office.context.document.settings.get(settingsKey);
        if (savedSpecs) {
            specs = JSON.parse(savedSpecs);
        }

        // Check if updating existing
        const specId = (document.getElementById("vega-spec-select") as HTMLSelectElement).value;
        
        if (specId) {
            const index = specs.findIndex((s: any) => s.id === specId);
            if (index > -1) {
                specs[index].name = chartName;
                specs[index].spec = spec;
            }
        } else {
            // Create new
            const newId = Date.now().toString();
            specs.push({ id: newId, name: chartName, spec: spec });
            
            // Update dropdown
            const select = document.getElementById("vega-spec-select") as HTMLSelectElement;
            const option = document.createElement("option");
            option.value = newId;
            option.text = chartName;
            select.appendChild(option);
            select.value = newId;
        }

        Office.context.document.settings.set(settingsKey, JSON.stringify(specs));
        Office.context.document.settings.saveAsync((result) => {
             if (result.status === Office.AsyncResultStatus.Failed) {
                 if (messageArea) messageArea.innerText = "Failed to save spec.";
             } else {
                 if (messageArea) messageArea.innerText = "Spec saved successfully.";
             }
        });

    } catch (e) {
        if (messageArea) messageArea.innerText = "Invalid JSON spec.";
    }
}

async function saveVegaChart() {
    showSpinner("Saving chart...");
    try {
    const messageArea = document.getElementById("vega-message");
    if (messageArea) messageArea.innerText = "Saving chart...";

    // First save spec
    await saveVegaSpec();
    
    try {
        const specStr = (document.getElementById("vega-spec") as HTMLTextAreaElement).value;
        const spec = JSON.parse(specStr);
        
        // Re-inject data just in case
        const tableName = (document.getElementById("vega-table-select") as HTMLSelectElement).value;
        const dt = await getArqueroTableFromExcel(tableName);
        const dataObjects = dt.objects();
        if (!spec.data || spec.data.name === "table") {
             spec.data = { values: dataObjects };
        }

        console.log("Generating view for image export...");
        const result = await vegaEmbed('#vega-preview-area', spec, { actions: false });
        const view = result.view;
        
        console.log("Converting to SVG...");
        const svg = await view.toSVG();
        // Encode SVG to base64, handling unicode characters correctly
        const base64 = btoa(unescape(encodeURIComponent(svg)));
        
        console.log("Inserting image into Excel...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const image = sheet.shapes.addImage(base64);
            image.name = (document.getElementById("vega-chart-name") as HTMLInputElement).value || "VegaChart";
            await context.sync();
        });
        
        if (messageArea) messageArea.innerText = "Chart inserted into worksheet.";
        
    } catch (e) {
        console.error("Save Chart Error:", e);
        if (messageArea) messageArea.innerText = "Error inserting chart: " + e.message;
    }
    } finally {
        hideSpinner();
    }
}