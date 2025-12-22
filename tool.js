#!/usr/bin/env node  
const fs = require("fs");  
const path = require("path");
  
function printUsage() {  
  console.log(`  
Usage: node script.js <command> [options]
  
Commands:  
  batch     Split a file into multiple batches  
  index     Add row indices to SQL statements  
  merge     Merge multiple files into one
  
Batch Options:  
  --input, -i <path>        Input file path (default: ./ROWS)  
  --amount, -a <number>     Rows per batch (default: 712)  
  --output, -o <dir>        Output directory (default: ./batches)  
  --table, -t <name>        Table name (default: [dbo].[tmdhtsAdditional])  
  --columns, -c <cols>      Comma-separated column names
  
Index Options:  
  --input, -i <path>        Input file path (default: ./ROWS)  
  --output, -o <path>       Output file path (default: ./ROWS_indexed)
  
Merge Options:  
  --folder, -f <path>       Folder containing files to merge (required)  
  --output, -o <path>       Output file path (default: <folder>_merged)
  
Examples:  
  node script.js batch -i ./data.sql -a 500 -o ./output  
  node script.js index -i ./ROWS -o ./ROWS_indexed  
  node script.js merge -f ./batches -o ./merged.sql  
`);  
}
  
function parseArgs(args) {  
  const params = {};  
  const positional = [];
  
  for (let i = 0; i < args.length; i++) {  
    const arg = args[i];
      
    if (arg.startsWith("--")) {  
      const key = arg.slice(2);  
      const value = args[i + 1];  
      params[key] = value;  
      i++;  
    } else if (arg.startsWith("-") && arg.length === 2) {  
      const flag = arg[1];  
      const value = args[i + 1];
        
      // Map short flags to long names  
      const flagMap = {  
        'i': 'input',  
        'a': 'amount',  
        'o': 'output',  
        't': 'table',  
        'c': 'columns',  
        'f': 'folder'  
      };
        
      if (flagMap[flag]) {  
        params[flagMap[flag]] = value;  
        i++;  
      }  
    } else {  
      positional.push(arg);  
    }  
  }
  
  return { params, positional };  
}
  
async function batcher(args) {  
  const { params } = parseArgs(args);
    
  const rowsPath = params.input || "./ROWS";  
  const amountPerBatch = parseInt(params.amount) || 712;  
  const outputDir = params.output || "./batches";  
  const tableName = params.table || "[dbo].[tmdhtsAdditional]";  
  const columns = params.columns || "[HTSNum], [Chapter99], [CountryofOrigin], [StartEffDate], [EndEffDate], [TariffType], [TariffGroup], [RequiredStatusCode], [ValidationLevel], [ExportDate]";
  
  console.log(`Reading file: ${rowsPath}`);  
  const file = fs.readFileSync(rowsPath, { encoding: "utf-8" });  
  const lines = file.toString().split("\n");  
  const totalBatches = Math.ceil(lines.length / amountPerBatch);
  
  console.log(`Total lines: ${lines.length}`);  
  console.log(`Batch size: ${amountPerBatch}`);  
  console.log(`Total batches: ${totalBatches}`);
  
  if (!fs.existsSync(outputDir)) {  
    fs.mkdirSync(outputDir, { recursive: true });  
  }
  
  for (let index = 0; index < totalBatches; index++) {  
    const batch = lines.splice(0, amountPerBatch);
      
    // Replace first line with INSERT statement  
    batch[0] = batch[0].replace(  
      ",(N'",  
      `INSERT INTO ${tableName} (${columns}) VALUES \n (N'`  
    );
      
    // Add semicolon to last line  
    const lastIndex = batch.length - 1;  
    batch[lastIndex] = `${batch[lastIndex].replace('\n', '').trimEnd()};`;
      
    const outputPath = path.join(outputDir, `batch_${index}.sql`);  
    fs.writeFileSync(outputPath, batch.join(""));  
    console.log(`Created batch ${index + 1}/${totalBatches}: ${outputPath}`);  
  }
    
  console.log(`Done! Created ${totalBatches} batches in ${outputDir}`);  
}
  
async function indexer(args) {  
  const { params } = parseArgs(args);
    
  const rowsPath = params.input || "./ROWS";  
  const outputPath = params.output || rowsPath + "_indexed";
  
  console.log(`Reading file: ${rowsPath}`);  
  const file = fs.readFileSync(rowsPath, { encoding: "utf-8" });  
  const lines = file.toString().split("\n");
    
  console.log(`Processing ${lines.length} lines...`);
    
  for (let index = 0; index < lines.length; index++) {  
    const line = lines[index];  
    lines[index] = line.replace(",(N'", `,(${index + 1}, N'`);  
  }
    
  fs.writeFileSync(outputPath, lines.join(''));  
  console.log(`Done! Indexed file written to: ${outputPath}`);  
}
  
async function merger(args) {  
  const { params } = parseArgs(args);
    
  const folder = params.folder;
    
  if (!folder) {  
    throw new Error("Folder path is required. Use --folder or -f");  
  }
    
  const outputPath = params.output || `${folder}_merged`;
  
  console.log(`Reading files from: ${folder}`);
    
  if (!fs.existsSync(folder)) {  
    throw new Error(`Folder does not exist: ${folder}`);  
  }
  
  const files = fs.readdirSync(folder).filter(f => {  
    const stat = fs.statSync(path.join(folder, f));  
    return stat.isFile();  
  });
    
  console.log(`Found ${files.length} files`);
    
  const contents = files.map((f) => {  
    const filePath = path.join(folder, f);  
    console.log(`Reading: ${filePath}`);  
    return fs.readFileSync(filePath).toString();  
  });
    
  fs.writeFileSync(outputPath, contents.join('\n'));  
  console.log(`Done! Merged file written to: ${outputPath}`);  
}
  
async function main() {  
  const args = process.argv.slice(2);
  
  if (args.length === 0 || args[0] === "--help" || args[0] === "-h") {  
    printUsage();  
    process.exit(0);  
  }
  
  const command = args[0];  
  const commandArgs = args.slice(1);
  
  try {  
    switch (command.toLowerCase()) {  
      case "merge":  
        await merger(commandArgs);  
        break;  
      case "batch":  
        await batcher(commandArgs);  
        break;  
      case "index":  
        await indexer(commandArgs);  
        break;  
      default:  
        console.error(`Unknown command: ${command}`);  
        printUsage();  
        process.exit(1);  
    }  
  } catch (error) {  
    console.error(`Error: ${error.message}`);  
    process.exit(1);  
  }  
}
  
main();  