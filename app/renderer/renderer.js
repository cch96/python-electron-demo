// import {PythonShell} from 'python-shell';
const {PythonShell} = require('python-shell')
firstBtn = document.querySelector('#firstTest')
secondBtn = document.querySelector('#secondTest')
firstBtn.onclick = () => {
    let options = {
        mode: 'text',
        pythonOptions: ['-u'], // get print results in real-time
        pythonPath: 'python37/Scripts/python.exe' 
    };
    PythonShell.run('core/code/FirstTest.py', options, function (err) {
        if (err) throw err;
        console.log('finished');
    });
}

secondBtn.onclick = () => {
    let options = {
        mode: 'text',
        pythonOptions: ['-u'], // get print results in real-time
        pythonPath: 'python37/Scripts/python.exe' 
    };
    PythonShell.run('core/code/secondTest.py', options, function (err) {
        if (err) throw err;
        console.log('finished');
    });
}