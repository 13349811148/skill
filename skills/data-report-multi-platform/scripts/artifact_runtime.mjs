// Resolve Codex's bundled runtime without a machine-specific node_modules link.
import {createRequire} from 'node:module';
import os from 'node:os';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
const local=createRequire(import.meta.url);
let entry;
try{entry=local.resolve('@oai/artifact-tool');}
catch{
 const root=process.env.CODEX_WORKSPACE_DEPENDENCIES||path.join(os.homedir(),'.cache/codex-runtimes/codex-primary-runtime/dependencies');
 const bundled=createRequire(path.join(root,'node','package.json'));
 try{entry=bundled.resolve('@oai/artifact-tool');}
 catch{throw new Error('Codex artifact runtime unavailable. Load workspace dependencies or set CODEX_WORKSPACE_DEPENDENCIES to the bundled dependencies directory.');}
}
const api=await import(pathToFileURL(entry).href);
export const {Workbook,SpreadsheetFile,FileBlob}=api;
