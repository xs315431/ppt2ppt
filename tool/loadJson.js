import path from "path"
import fs from "fs"
import { fileURLToPath } from 'url'; 

export function loadJson(url) {

    // 获取当前文件路径
    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);

    // 读取 JSON
    const jsonPath = path.resolve(__dirname, url);
    const rawData = fs.readFileSync(jsonPath, 'utf8');
    const jsonData = JSON.parse(rawData)
    return jsonData
}

