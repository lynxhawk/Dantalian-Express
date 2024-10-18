const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;
const cors = require('cors');
app.use(cors());


// 设置解析 JSON 的中间件
app.use(express.json());

// Excel 文件路径（固定路径）
const EXCEL_FILE_PATH = path.join(__dirname, 'BookInfo.xlsx');

// 检查并创建初始 Excel 文件
const initializeExcelFile = () => {
  if (!fs.existsSync(EXCEL_FILE_PATH)) {
    // 如果文件不存在，则创建一个新的文件并添加初始工作表
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['Title', 'Authors', 'Publisher', 'Published Date', 'Description', 'Page Count', 'Categories', 'Language', 'Thumbnail'],
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Books');
    XLSX.writeFile(workbook, EXCEL_FILE_PATH);
  }
};

// 处理读取并附加新图书内容的请求
app.post('/add-book', (req, res) => {
  try {
    // 初始化 Excel 文件（如果不存在）
    initializeExcelFile();

    // 读取现有的 Excel 文件
    const workbook = XLSX.readFile(EXCEL_FILE_PATH);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // 获取请求体中的新书籍数据
    const newBook = req.body;

    // 检查请求体中的必要字段
    if (!newBook.title || !newBook.authors) {
      return res.status(400).send('Title and Authors are required.');
    }

    // 将新书籍信息附加到工作表中
    const newData = [
      {
        Title: newBook.title,
        Authors: Array.isArray(newBook.authors) ? newBook.authors.join(', ') : newBook.authors,
        Publisher: newBook.publisher || '未知',
        'Published Date': newBook.publishedDate || '未知',
        Description: newBook.description || '暂无描述',
        'Page Count': newBook.pageCount || '未知',
        Categories: newBook.categories ? newBook.categories.join(', ') : '未知',
        Language: newBook.language || '未知',
        Thumbnail: newBook.thumbnail || '无',
      },
    ];

    // 将新数据追加到工作表中
    XLSX.utils.sheet_add_json(worksheet, newData, { skipHeader: true, origin: -1 });

    // 将更新后的工作表写入到文件中
    XLSX.writeFile(workbook, EXCEL_FILE_PATH);

    res.status(200).send('新书籍信息已成功添加到 Excel 文件中');
  } catch (error) {
    console.error('发生错误:', error);
    res.status(500).send('服务器错误');
  }
});

// 处理获取所有书籍信息的请求
app.get('/books', (req, res) => {
  try {
    // 初始化 Excel 文件（如果不存在）
    initializeExcelFile();

    // 读取现有的 Excel 文件
    const workbook = XLSX.readFile(EXCEL_FILE_PATH);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // 将工作表内容转换为 JSON
    const books = XLSX.utils.sheet_to_json(worksheet);

    res.status(200).json(books);
  } catch (error) {
    console.error('发生错误:', error);
    res.status(500).send('服务器错误');
  }
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`服务器正在 http://localhost:${PORT} 上运行`);
});