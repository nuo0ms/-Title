const BASE = 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
let py;
let chain = Promise.resolve();
async function handle({id,type,payload}) {
  try {
    let result = '{}';
    if(type==='boot') {
      importScripts(BASE+'pyodide.js');
      py=await loadPyodide({indexURL:BASE});
      await py.loadPackage(['numpy','micropip']);
      await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl==3.1.5')");
      const response=await fetch('./model.py');
      if(!response.ok) throw Error('model.py를 불러올 수 없습니다.');
      await py.runPythonAsync(await response.text());
    } else if(type==='load' || type==='classify') {
      py.globals.set('raw_input',new Uint8Array(payload.buffer));
      py.globals.set('level_input',payload.level||'');
      await py.runPythonAsync('raw_bytes = raw_input.to_py()');
      result=await py.runPythonAsync(type==='load'?'initialize(level_input, raw_bytes)':'load_class(raw_bytes)');
      await py.runPythonAsync('del raw_input, raw_bytes');
    } else if(type==='explore') {
      py.globals.set('level_input',payload.level);
      py.globals.set('sheet_input',payload.sheet);
      py.globals.set('row_input',payload.row||'');
      result=py.runPython('explore(level_input, sheet_input, row_input)');
    } else if(type==='linkage') {
      py.globals.set('level_input',payload.level);
      result=py.runPython('linkage(level_input)');
    } else if(type==='items') {
      py.globals.set('level_input',payload.level);
      py.globals.set('code_input',payload.code);
      result=py.runPython('demand_items(level_input, code_input)');
    } else {
      py.globals.set('level_input',payload.level);
      py.globals.set('scenario_input',JSON.stringify(payload.shocks));
      result=py.runPython('simulate(level_input, scenario_input)');
    }
    postMessage({id,ok:true,result:JSON.parse(result)});
  } catch(error) { postMessage({id,ok:false,error:String(error)}); }
}
onmessage=e=>{ chain=chain.then(()=>handle(e.data)); };
