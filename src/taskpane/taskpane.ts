/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import * as aq from 'arquero';
import vegaEmbed from 'vega-embed';
import mermaid from 'mermaid';

// Global state for Vega auto-update
let vegaAutoUpdateHandler: OfficeExtension.EventHandlerResult<Excel.TableChangedEventArgs> = null;
let vegaCurrentTableName: string = null;
let vegaCurrentSpec: any = null;

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

    document.getElementById("close-arquero").onclick = () => {
        editor.style.display = "none";
    };
    
    document.getElementById("run-preview").onclick = runPreview;
    document.getElementById("save-output").onclick = outputToExcel;
}

async function getArqueroTableFromExcel(tableName: string) {
    let arqueroTable;
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(tableName);
        const headerRange = table.getHeaderRowRange();
        const dataRange = table.getDataBodyRange();
        
        headerRange.load("values");
        dataRange.load("values");
        
        // Single sync for both header and data
        await context.sync();

        const headers = headerRange.values[0];
        const rows = dataRange.values;

        // Use columnar format for Arquero (much faster than row-wise objects)
        const columnData = {};
        headers.forEach((header, colIndex) => {
            columnData[header] = rows.map(row => row[colIndex]);
        });
        
        arqueroTable = aq.table(columnData);
    });
    return arqueroTable;
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
    let tableName: string = null;
    
    // Try to get the active table from the current selection
    await Excel.run(async (context) => {
        try {
            const activeCell = context.workbook.getActiveCell();
            const tables = activeCell.getTables(false);
            tables.load("items/name");
            await context.sync();
            
            if (tables.items.length > 0) {
                tableName = tables.items[0].name;
            }
        } catch (error) {
            console.error("Error getting active table:", error);
        }
    });
    
    if (!tableName) {
        throw new Error("No table selected. Please click on a cell within an Excel table and try again.");
    }
    
    return await getArqueroTableFromExcel(tableName);
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

async function closeVegaEditor() {
    // Clean up auto-update event handler
    if (vegaAutoUpdateHandler) {
        await Excel.run(async (context) => {
            vegaAutoUpdateHandler.remove();
            await context.sync();
        }).catch(err => console.error("Error removing event handler:", err));
        vegaAutoUpdateHandler = null;
        vegaCurrentTableName = null;
        vegaCurrentSpec = null;
    }
    
    document.getElementById("vega-editor").style.display = "none";
    
    // Clear status indicator
    const statusDiv = document.getElementById("vega-auto-update-status");
    if (statusDiv) statusDiv.style.display = "none";
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
        
        // Store current spec for auto-update
        vegaCurrentSpec = spec;
        vegaCurrentTableName = tableName;
        
        // Render chart with current data
        await renderVegaChart(tableName, spec, messageArea);
        
        // Set up auto-update if table changed or no handler exists
        await setupVegaAutoUpdate(tableName, messageArea);

    } catch (e) {
        console.error("Vega Preview Error:", e);
        if (messageArea) messageArea.innerText = "Error rendering chart: " + e.message;
    }
    } finally {
        hideSpinner();
    }
}

async function renderVegaChart(tableName: string, spec: any, messageArea?: HTMLElement) {
    console.log(`Fetching data for table: ${tableName}`);
    // Get Data
    const dt = await getArqueroTableFromExcel(tableName);
    const dataObjects = dt.objects();
    console.log(`Data fetched. Rows: ${dataObjects.length}`);
    
    // Remove $schema to avoid CSP violations in Excel Online
    delete spec.$schema;
    
    // Inject data if not present or if placeholder used
    if (!spec.data || spec.data.name === "table") {
         spec.data = { values: dataObjects };
    }

    console.log("Embedding chart...");
    // Render
    await vegaEmbed('#vega-preview-area', spec, { actions: false });
    console.log("Chart rendered.");
    if (messageArea) messageArea.innerText = "";
}

async function setupVegaAutoUpdate(tableName: string, messageArea?: HTMLElement) {
    // Remove existing handler if table changed
    if (vegaAutoUpdateHandler && vegaCurrentTableName !== tableName) {
        await Excel.run(async (context) => {
            vegaAutoUpdateHandler.remove();
            await context.sync();
        }).catch(err => console.error("Error removing old handler:", err));
        vegaAutoUpdateHandler = null;
    }
    
    // Add new handler if needed
    if (!vegaAutoUpdateHandler) {
        await Excel.run(async (context) => {
            const table = context.workbook.tables.getItem(tableName);
            
            vegaAutoUpdateHandler = table.onChanged.add(async (event) => {
                console.log(`Table '${tableName}' changed, auto-updating chart...`);
                try {
                    await renderVegaChart(vegaCurrentTableName, vegaCurrentSpec, messageArea);
                    console.log("Chart auto-updated successfully.");
                } catch (err) {
                    console.error("Error auto-updating chart:", err);
                }
            });
            
            await context.sync();
            console.log(`Auto-update enabled for table '${tableName}'`);
            
            // Show status indicator
            const statusDiv = document.getElementById("vega-auto-update-status");
            if (statusDiv) {
                statusDiv.innerText = `🔄 Auto-update active for table: ${tableName}`;
                statusDiv.style.display = "block";
            }
        });
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
        
        console.log("Converting SVG to PNG...");
        // Create an image element to load the SVG
        const img = new Image();
        const svgBlob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
        const svgUrl = URL.createObjectURL(svgBlob);
        
        // Wait for image to load
        await new Promise((resolve, reject) => {
            img.onload = resolve;
            img.onerror = reject;
            img.src = svgUrl;
        });
        
        // Create canvas and draw the image
        const canvas = document.createElement('canvas');
        canvas.width = img.width || 800;
        canvas.height = img.height || 600;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0);
        
        // Clean up
        URL.revokeObjectURL(svgUrl);
        
        // Convert canvas to PNG data URL
        const pngDataUri = canvas.toDataURL('image/png');
        console.log("PNG data URI created, length:", pngDataUri.length);
        console.log("PNG data URI prefix:", pngDataUri.substring(0, 50));
        
        console.log("Inserting image into Excel...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            console.log("Got active worksheet");
            
            const imageName = (document.getElementById("mermaid-chart-name") as HTMLInputElement).value || "MermaidDiagram";
            console.log("Calling addImage with name:", imageName);
            
            const image = sheet.shapes.addImage(pngDataUri);
            image.name = imageName;
            console.log("Image shape created, calling sync...");
            
            await context.sync();
            console.log("Sync completed successfully");
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
        
        // Remove $schema to avoid CSP violations in Excel Online
        delete spec.$schema;
        
        // Re-inject data just in case
        const tableName = (document.getElementById("vega-table-select") as HTMLSelectElement).value;
        const dt = await getArqueroTableFromExcel(tableName);
        const dataObjects = dt.objects();
        if (!spec.data || spec.data.name === "table") {
             spec.data = { values: dataObjects };
        }

        console.log("Generating view for image export...");
        const result = await vegaEmbed('#vega-preview-area', spec, { actions: false, loader: { http: { credentials: 'same-origin' } } });
        const view = result.view;
        
        console.log("Converting to PNG...");
        const canvas = await view.toCanvas();
        console.log("Canvas created, width:", canvas.width, "height:", canvas.height);
        
        // Convert canvas to PNG data URL
        const pngDataUri = canvas.toDataURL('image/png');
        console.log("PNG data URI created, length:", pngDataUri.length);
        console.log("PNG data URI prefix:", pngDataUri.substring(0, 50));
        
        console.log("Inserting image into Excel...");
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            console.log("Got active worksheet");
            
            const imageName = (document.getElementById("vega-chart-name") as HTMLInputElement).value || "VegaChart";
            console.log("Calling addImage with name:", imageName);
            
            const image = sheet.shapes.addImage(pngDataUri);
            image.name = imageName;
            console.log("Image shape created, calling sync...");
            
            await context.sync();
            console.log("Sync completed successfully");
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