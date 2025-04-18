const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');
const inquirer = require('inquirer'); 
// 检查链接是否可以访问
async function checkLinkAccessibility(url) {
    try {
        const response = await axios.head(url, { timeout: 5000 }); // 使用 HEAD 请求快速检查
        return response.status >= 200 && response.status < 400; // 状态码在 200-399 范围内视为有效
    } catch (error) {
        return false; // 如果发生错误（如超时或无效链接），返回 false
    }
}

// 读取 Excel 文件并检测链接
async function detectBrokenLinksInExcel(filePath) {
    // 读取 Excel 文件
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // 假设我们只处理第一个工作表
    const worksheet = workbook.Sheets[sheetName];

    // 将工作表转换为 JSON 格式
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    console.log('开始检测链接...');
    const brokenLinks = [];

    for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
        const row = data[rowIndex];
        for (let colIndex = 0; colIndex < row.length; colIndex++) {
            const cellValue = row[colIndex];
            if (typeof cellValue === 'string' && isValidUrl(cellValue)) {
                const isAccessible = await checkLinkAccessibility(cellValue);
                if (!isAccessible) {
                    brokenLinks.push({
                        link: cellValue,
                        location: `Row ${rowIndex + 1}, Column ${colIndex + 1}`
                    });
                }
            }
        }
    }

    if (brokenLinks.length > 0) {
        console.log('发现以下无法访问的链接：');
        brokenLinks.forEach((item, index) => {
            console.log(`${index + 1}. 链接: ${item.link} | 位置: ${item.location}`);
        });
    } else {
        console.log('所有链接均可访问！');
    }
}

// 检查字符串是否为有效的 URL
function isValidUrl(string) {
    try {
        new URL(string);
        return true;
    } catch (_) {
        return false;
    }
}

// 列出当前目录下的所有 .xlsx 文件
function listExcelFiles() {
    const files = fs.readdirSync(process.cwd()) // 获取当前目录下的所有文件
        .filter(file => path.extname(file).toLowerCase() === '.xlsx'); // 筛选出 .xlsx 文件
    return files;
}

// 主函数：交互式选择文件
(async () => {
    const excelFiles = listExcelFiles();

    if (excelFiles.length === 0) {
        console.log('当前目录下没有找到任何 .xlsx 文件。');
        return;
    }

    console.log('当前目录下的 .xlsx 文件列表：');
    excelFiles.forEach((file, index) => {
        console.log(`${index + 1}. ${file}`);
    });

    const questions = [
        {
            type: 'input',
            name: 'fileIndex',
            message: '请输入文件序号选择 Excel 文件：',
            validate: (input) => {
                const index = parseInt(input, 10);
                if (isNaN(index) || index < 1 || index > excelFiles.length) {
                    return `请输入有效的序号（1-${excelFiles.length}）。`;
                }
                return true;
            }
        }
    ];

    try {
        const answers = await inquirer.prompt(questions); // 确保 inquirer 是最新版本
        const selectedFile = excelFiles[parseInt(answers.fileIndex, 10) - 1];
        const filePath = path.resolve(selectedFile); // 获取绝对路径

        console.log(`已选择文件: ${selectedFile}`);
        await detectBrokenLinksInExcel(filePath);
    } catch (err) {
        console.error('发生错误:', err);
    }
})();