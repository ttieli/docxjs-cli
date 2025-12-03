const fs = require('fs');
const path = require('path');

const TEMPLATES_DIR = path.join(__dirname, '..', 'templates');
const BUILTIN_TEMPLATE_FILES = ['templates.json', 'common_styles.json']; // 这些文件里的模板是内置的

class TemplateManager {
    constructor() {
        this.templates = this._loadAllTemplates();
    }

    _loadAllTemplates() {
        let allTemplates = {};
        if (!fs.existsSync(TEMPLATES_DIR)) {
            console.warn(`⚠️ Templates directory not found at: ${TEMPLATES_DIR}`);
            return {};
        }

        try {
            const files = fs.readdirSync(TEMPLATES_DIR).filter(file => file.toLowerCase().endsWith('.json'));
            files.sort(); // 确保加载顺序一致

            files.forEach(file => {
                const fullPath = path.join(TEMPLATES_DIR, file);
                try {
                    const fileContent = JSON.parse(fs.readFileSync(fullPath, 'utf-8'));
                    // 将每个模板标记为来自哪个文件，以便后续判断是否内置
                    Object.keys(fileContent).forEach(key => {
                        fileContent[key].__sourceFile = file;
                    });
                    allTemplates = { ...allTemplates, ...fileContent };
                } catch (e) {
                    console.warn(`⚠️ Warning: Failed to parse template file '${file}': ${e.message}`);
                }
            });
        } catch (e) {
            console.error(`❌ Error reading templates directory: ${e.message}`);
        }
        return allTemplates;
    }

    _getTemplateFilePath(templateName) {
        // 如果是内置模板，返回其源文件路径
        const template = this.templates[templateName];
        if (template && template.__sourceFile) {
            return path.join(TEMPLATES_DIR, template.__sourceFile);
        }
        // 对于用户新增的模板，我们假设保存到 user_templates.json
        return path.join(TEMPLATES_DIR, 'user_templates.json');
    }

    // 检查是否为内置模板
    isBuiltIn(templateName) {
        const template = this.templates[templateName];
        if (!template || !template.__sourceFile) return false;
        return BUILTIN_TEMPLATE_FILES.includes(template.__sourceFile);
    }

    // 获取所有模板 (Read All)
    getAllTemplates() {
        // 返回不包含内部 __sourceFile 属性的模板列表
        return Object.entries(this.templates).map(([key, value]) => {
            const { __sourceFile, ...rest } = value;
            return { name: key, ...rest, isBuiltIn: this.isBuiltIn(key) };
        });
    }

    // 获取单个模板 (Read One)
    getTemplate(templateName) {
        const template = this.templates[templateName];
        if (!template) return null;
        const { __sourceFile, ...rest } = template;
        return { name: templateName, ...rest, isBuiltIn: this.isBuiltIn(templateName) };
    }

    // 保存/更新模板 (Create/Update)
    // 假设用户自定义模板都保存到 user_templates.json
    saveTemplate(templateName, config) {
        if (!templateName || !config) {
            throw new Error("Template name and config are required.");
        }
        if (this.isBuiltIn(templateName)) {
            throw new Error(`Cannot modify built-in template: ${templateName}`);
        }

        const userTemplatesFilePath = this._getTemplateFilePath(templateName); // Should be user_templates.json
        let userTemplates = {};
        if (fs.existsSync(userTemplatesFilePath)) {
            try {
                userTemplates = JSON.parse(fs.readFileSync(userTemplatesFilePath, 'utf-8'));
            } catch (e) {
                console.warn(`⚠️ Could not parse existing user_templates.json, starting fresh.`);
            }
        }
        
        userTemplates[templateName] = { ...config, __sourceFile: 'user_templates.json' };
        fs.writeFileSync(userTemplatesFilePath, JSON.stringify(userTemplates, null, 2));
        this.templates = this._loadAllTemplates(); // 重新加载所有模板
        return this.getTemplate(templateName);
    }

    // 删除模板 (Delete)
    deleteTemplate(templateName) {
        if (this.isBuiltIn(templateName)) {
            throw new Error(`Cannot delete built-in template: ${templateName}`);
        }
        if (!this.templates[templateName]) {
            throw new Error(`Template not found: ${templateName}`);
        }

        const userTemplatesFilePath = this._getTemplateFilePath(templateName); // Should be user_templates.json
        if (!fs.existsSync(userTemplatesFilePath)) {
             throw new Error(`User template file not found for ${templateName}`);
        }
        let userTemplates = JSON.parse(fs.readFileSync(userTemplatesFilePath, 'utf-8'));
        
        if (userTemplates[templateName]) {
            delete userTemplates[templateName];
            fs.writeFileSync(userTemplatesFilePath, JSON.stringify(userTemplates, null, 2));
            this.templates = this._loadAllTemplates(); // 重新加载所有模板
            return true;
        }
        return false;
    }
}

module.exports = new TemplateManager(); // 导出单例
