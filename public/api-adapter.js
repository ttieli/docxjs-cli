const API = {
    isElectron: () => !!window.electronAPI,
  
    async getTemplates() {
      if (this.isElectron()) {
        return await window.electronAPI.getTemplates();
      } else {
        const res = await fetch('/api/templates');
        if (!res.ok) throw new Error('Failed to fetch templates');
        return await res.json();
      }
    },
  
    async importDocx(file) {
      if (this.isElectron()) {
        // Convert File to ArrayBuffer
        const buffer = await file.arrayBuffer();
        return await window.electronAPI.importDocx(buffer);
      } else {
        const formData = new FormData();
        formData.append('file', file);
        const res = await fetch('/api/import-docx', { method: 'POST', body: formData });
        if (!res.ok) throw new Error('Word parsing failed');
        return await res.json();
      }
    },
  
    async convert(params) {
      // params: { markdown, templateName, filename, referenceDoc (File), styleConfig (Obj) }
      if (this.isElectron()) {
          const data = {
              markdown: params.markdown,
              templateName: params.templateName,
              filename: params.filename,
              styleConfig: params.styleConfig
          };
          if (params.referenceDoc) {
              data.referenceDocBuffer = await params.referenceDoc.arrayBuffer();
          }
          // Returns Uint8Array
          const buffer = await window.electronAPI.convert(data);
          // Convert back to Blob for download
          return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      } else {
        const formData = new FormData();
        formData.append('markdown', params.markdown);
        if (params.templateName) formData.append('templateName', params.templateName);
        if (params.filename) formData.append('filename', params.filename);
        if (params.referenceDoc) formData.append('referenceDoc', params.referenceDoc);
        if (params.styleConfig) formData.append('styleConfig', JSON.stringify(params.styleConfig));
  
        const res = await fetch('/api/convert', { method: 'POST', body: formData });
        if (!res.ok) {
          const errText = await res.text();
          let msg = 'Conversion failed';
          try { msg = JSON.parse(errText).error || msg; } catch (_) { msg = errText || msg; }
          throw new Error(msg);
        }
        return await res.blob();
      }
    },

    async preview(styleConfig) {
      if (this.isElectron()) {
        const buffer = await window.electronAPI.previewTemplateDynamic(styleConfig);
        return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      } else {
        const res = await fetch('/api/preview/dynamic', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ styleConfig })
        });
        if (!res.ok) throw new Error('Preview failed');
        return await res.blob();
      }
    }
  };
