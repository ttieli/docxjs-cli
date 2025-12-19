const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    getTemplates: () => ipcRenderer.invoke('get-templates'),
    importDocx: (fileBuffer) => ipcRenderer.invoke('import-docx', fileBuffer),
    convert: (data) => ipcRenderer.invoke('convert', data),
    previewTemplateDynamic: (styleConfig, markdown) => ipcRenderer.invoke('preview-template-dynamic', { styleConfig, markdown }),
    printToPDF: () => ipcRenderer.invoke('print-to-pdf')
});
