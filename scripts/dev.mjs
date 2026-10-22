import { spawn } from 'node:child_process';
const commands = [['node',['--env-file-if-exists=.env','--watch','server/src/index.js']],['node',['node_modules/vite/bin/vite.js','--host','127.0.0.1']]];
const children = commands.map(([cmd,args]) => spawn(cmd,args,{stdio:'inherit'}));
let closing=false;
function stop(code=0){ if(closing)return;closing=true;children.forEach(c=>c.kill('SIGTERM'));process.exitCode=code; }
process.on('SIGINT',()=>stop());process.on('SIGTERM',()=>stop());
children.forEach(c=>c.on('exit',code=>stop(code||0)));
