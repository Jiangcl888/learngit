const fs = require('fs');
const path = require('path');
const axios = require('axios');
const inquirer = require('inquirer');
const ExcelJS = require('exceljs');

// 检查链接是否可以访问
async function checkLinkAccessibility(url) {
    try {
        const response = await axios.head(url, { timeout: 5000 }); // 使用 HEAD 请求快速检查
        return response.status >= 200 && response.status < 400; // 状态码在 200-399 范围内视为有效
    } catch (error) {
        return false; // 如果发生错误（如超时或无效链接），返回 false
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

// 主函数：交互式选择文件并处理
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
        const answers = await inquirer.prompt(questions);
        const selectedFile = excelFiles[parseInt(answers.fileIndex, 10) - 1];
        const filePath = path.resolve(selectedFile); // 获取绝对路径

        console.log(`已选择文件: ${selectedFile}`);
        await processExcelFile(filePath);
    } catch (err) {
        console.error('发生错误:', err);
    }
})();

// 处理 Excel 文件并标记错误链接
async function processExcelFile(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.worksheets[0]; // 假设我们只处理第一个工作表
    const brokenLinks = [];

    // 遍历每个单元格
    worksheet.eachRow((row, rowIndex) => {
        row.eachCell((cell, colIndex) => {
            if (typeof cell.value === 'string' && isValidUrl(cell.value)) {
                const url = cell.value;
                brokenLinks.push({
                    url,
                    location: { row: rowIndex, col: colIndex + 1 }, // 单元格位置
                    cell
                });
            }
        });
    });

    // 检测每个链接的状态
    for (const link of brokenLinks) {
        const isAccessible = await checkLinkAccessibility(link.url);
        if (!isAccessible) {
            console.log(`发现错误链接: ${link.url} | 位置: Row ${link.location.row}, Column ${link.location.col}`);
            // 设置字体颜色为红色
            link.cell.font = {
                color: { argb: 'FFFF0000' }, // 设置字体颜色为红色
                bold: true // 可选：加粗字体
            };
        }
    }

    if (brokenLinks.length > 0) {
        // 保存修改后的 Excel 文件
        const outputFilePath = path.join(path.dirname(filePath), `marked_${path.basename(filePath)}`);
        await workbook.xlsx.writeFile(outputFilePath);
        console.log(`已将错误链接标记为红色，并保存到新文件: ${outputFilePath}`);
    } else {
        console.log('所有链接均可访问！');
    }
}