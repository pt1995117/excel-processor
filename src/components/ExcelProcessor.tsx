import React, { useState } from 'react';
import { Upload, Button, message, Table, Spin, Input } from 'antd';
import { UploadOutlined, LineChartOutlined } from '@ant-design/icons';
import * as XLSX from 'xlsx';
import type { UploadFile } from 'antd/es/upload/interface';
import axios from 'axios';

/**
 * @interface ProcessedTable
 * @description 处理后的表格数据接口
 */
interface ProcessedTable {
  name: string;
  data: any[];
  columns: any[];
  analysis?: string;
  isAnalyzing: boolean;
  commonTopics: string[];
  isAnalyzingTopics: boolean;
  topicsAnalysis?: string; // 新增：共性内容分析结果
  topicsApiStatus?: 'success' | 'error' | 'analyzing' | null;
  topicsApiMessage?: string;
  analysisApiStatus?: 'success' | 'error' | 'analyzing' | null;
  analysisApiMessage?: string;
}

/**
 * @interface ExcelData
 * @description Excel数据接口
 */
interface ExcelData {
  [key: string]: string | number;
}

// DeepSeek API配置
// const DEEPSEEK_API_KEY = process.env.REACT_APP_DEEPSEEK_API_KEY || '';
// const DEEPSEEK_API_URL = process.env.REACT_APP_DEEPSEEK_API_URL || 'https://api.deepseek.com/v1/chat/completions';

// API配置
const ARK_API_KEY = process.env.REACT_APP_ARK_API_KEY || '';
const ARK_API_URL = 'https://ark.cn-beijing.volces.com/api/v3/chat/completions';
const MODEL_NAME = process.env.REACT_APP_MODEL_NAME || 'doubao-1-5-pro-256k-250115';

/**
 * @component ExcelProcessor
 * @description 问卷数据分析组件
 */
const ExcelProcessor: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [tables, setTables] = useState<ProcessedTable[]>([]);
  const [commonTopicsInput, setCommonTopicsInput] = useState('');

  /**
   * @function shouldSkipColumn
   * @description 判断是否应该跳过某列 - 按需求文档要求保留包含"text"或"分列"字样的表头
   * @param {string} columnName - 列名
   * @returns {boolean}
   */
  const shouldSkipColumn = (columnName: string): boolean => {
    // 按需求文档：保留包含"text"或"分列"字样的表头，其余舍弃
    const lowerColumnName = columnName.toLowerCase();
    return !lowerColumnName.includes('text') && !lowerColumnName.includes('分列');
  };

  /**
   * @function hasEnoughUniqueValues
   * @description 检查列是否有足够的非空唯一值 - 按需求文档要求处理空值和"无"，去重后至少10条
   * @param {ExcelData[]} data - Excel数据
   * @param {string} columnName - 列名
   * @returns {boolean}
   */
  const hasEnoughUniqueValues = (data: ExcelData[], columnName: string): boolean => {
    // 按需求文档：删除最后一列为空值或"无"的整行后，检查是否还有数据
    const validRows = data.filter(row => {
      const value = row[columnName];
      return value !== undefined && value !== null && value !== '' && value !== '无';
    });
    
    // 去重：提取有效值并去重
    const uniqueValues = new Set(
      validRows.map(row => {
        const value = row[columnName];
        // 转换为字符串并去除首尾空格进行比较
        return typeof value === 'string' ? value.trim() : String(value);
      }).filter(value => value !== '')
    );
    
    // 按需求文档：去重后至少要有10条有效数据
    const uniqueCount = uniqueValues.size;
    console.log(`列 "${columnName}" 去重后有效数据条数: ${uniqueCount}`);
    
    return uniqueCount >= 10;
  };

  /**
   * @function handleAnalyze
   * @description 处理单个表格的分析请求 - 按需求文档要求实现分批处理和汇总
   * @param {number} index - 表格索引
   */
  const handleAnalyze = async (index: number) => {
    if (tables[index].isAnalyzing) return;

    // 保存当前表格的引用，避免在处理过程中状态变化导致的问题
    const currentTable = tables[index];

    // 更新状态：开始分析
    setTables(prev => prev.map((table, i) => 
      i === index ? { 
        ...table, 
        isAnalyzing: true,
        analysisApiStatus: 'analyzing',
        analysisApiMessage: '正在进行整体分析...'
      } : table
    ));

    try {
      const BATCH_SIZE = 300;
      const batches = [];
      
      // 创建批次
      for (let i = 0; i < currentTable.data.length; i += BATCH_SIZE) {
        batches.push(currentTable.data.slice(i, i + BATCH_SIZE));
      }

      // 更新API状态：开始批量处理
      setTables(prev => prev.map((t, i) => 
        i === index ? { 
          ...t, 
          analysisApiMessage: `开始批量处理，共${batches.length}批，每批最多${BATCH_SIZE}条数据`
        } : t
      ));

      const batchResults: string[] = [];

      // 处理每个批次
      for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
        const batch = batches[batchIndex];
        
        // 更新进度
        setTables(prev => prev.map((t, i) => 
          i === index ? { 
            ...t, 
            analysisApiMessage: `正在处理第${batchIndex + 1}/${batches.length}批数据...`
          } : t
        ));

        // 从表格名称中提取实际的列名
        const columnName = currentTable.name.replace(/^表格 - /, '').replace(/（共 \d+ 条数据）$/, '');
        
        const batchData = batch.map(row => ({
          用户信息: `${row['姓名'] || ''}-${row['系统号'] || ''}`,
          问题内容: row[columnName] || row[currentTable.name] || ''
        }));

        const batchPrompt = `# 角色
你是一位资深的数据处理分析师，在数据清洗与预处理领域拥有深厚造诣，同时具备卓越的基于数据进行深入文本分析的能力。你以严谨、细致、高效的态度完成各项数据处理与分析任务。

## 技能
### 技能 1: 数据清洗与预处理
1. 接收输入数据后，运用专业方法全方位排查数据中存在的各类问题，涵盖缺失值、重复值、异常值等。深入剖析问题产生的原因，不局限于问题的发现。
2. 依据数据的特征、来源以及后续分析需求，精准挑选最合适的方法对这些问题进行妥善处理，保证数据质量达到高质量分析所需的严格标准。处理过程需详细记录文档，以便追溯。

### 技能 2: 文本分析与共性问题识别
1. 不管输入数据是否规范，针对表头的问题内容以及用户提供的回答展开深度文本分析。不仅进行词频统计、语义分析等常规操作，还运用关联分析等技术手段挖掘潜在文本关系。
2. 通过多种技术手段，精确识别其中的共性问题。共性问题提取要具体准确，避免过于宽泛的概括（例如不能简单概括为管理），对于难以准确归类的问题，单独列出说明。
3. 统计每个共性问题被提及的用户数量，详细罗列提出共性问题的所有用户姓名以及用户系统号，确保无遗漏。同时，记录每个用户提出问题的时间信息。
4. 针对共性问题下高频出现的内容进行系统总结，统计每个高频内容被提及的人数。除人数统计外，分析高频内容出现的趋势变化。
5. 从每个高频内容中选取有代表性的原文内容呈现出来，并备注输出该原文内容的用户姓名、系统号以及相关时间信息。

## 限制:
- 回答必须紧密围绕数据清洗、预处理以及文本分析相关任务，坚决拒绝回答无关话题。
- 输出内容要以清晰、符合逻辑的格式呈现共性问题、用户姓名、用户系统号、共性问题下高频出现内容的总结、全部被提及人数以及相关时间信息，不得有任何遗漏。 
- 千万不要遗漏提及共性问题的用户统计与输出，包括用户提出问题的时间信息。
注意：
- 一级标题包括：提及人数>=3共性内容 + 提及人数>=3的非共性内容
- 按提及人数从多到少排序，若人数相同，按照问题首次出现的时间先后排序
- 每个标题下必须包含用户清单和具体问题描述
- 负向内容必须全部呈现，不论提及人数多少
- 每个具体问题都要标注提及人数并引用代表性原文
- 不得因篇幅过长而省略任何问题或内容
- 所有满足条件的问题必须完整展示，包括所有用户信息和原文引用
- 确保输出所有相关内容，即使最终分析结果较长，要注重数据的完整性和准确性。 

## 输出格式：
[一级标题]（XX人提及）
- 提出用户：
  * 姓名：xxx，系统号：xxx
  * 姓名：xxx，系统号：xxx
- 具体问题：
  1. 问题1（XX人提及）：
     - 问题描述：xxx
     - 代表性原文："xxx"（来自：姓名-系统号）
     - 代表性原文："xxx"（来自：姓名-系统号）
  2. 问题2（XX人提及）：
     - 问题描述：xxx
     - 代表性原文："xxx"（来自：姓名-系统号）
     - 代表性原文："xxx"（来自：姓名-系统号）

[负向反馈]
- 问题1（XX人提及）：
  * 问题描述：xxx
  * 提出用户：姓名-系统号
  * 原文内容："xxx"
- 问题2（XX人提及）：
  * 问题描述：xxx
  * 提出用户：姓名-系统号
  * 原文内容："xxx" 

请分析以下数据：
${JSON.stringify(batchData, null, 2)}`;

        try {
          const response = await fetch('https://ark.cn-beijing.volces.com/api/v3/chat/completions', {
            method: 'POST',
          headers: {
            'Content-Type': 'application/json',
              'Authorization': 'Bearer af954c6d-7c97-4d29-926e-873807ed6032'
            },
            body: JSON.stringify({
              model: 'doubao-1-5-pro-256k-250115',
              messages: [
                { role: 'system', content: '你是一个数据分析专家，擅长文本分析和共性问题识别。' },
                { role: 'user', content: batchPrompt }
              ]
            })
          });

          if (!response.ok) {
            throw new Error(`批次${batchIndex + 1}API请求失败: ${response.status} ${response.statusText}`);
          }

          const result = await response.json();
          
          if (!result.choices || !result.choices[0] || !result.choices[0].message) {
            throw new Error(`批次${batchIndex + 1}API返回数据格式错误`);
          }

          batchResults.push(result.choices[0].message.content);
          console.log(`✓ 批次${batchIndex + 1}分析完成，结果长度：${result.choices[0].message.content.length}`);
          
        } catch (batchError) {
          console.error(`批次${batchIndex + 1}处理失败:`, batchError);
          batchResults.push(`批次${batchIndex + 1}分析失败: ${batchError instanceof Error ? batchError.message : '未知错误'}`);
        }
      }

      // 更新API状态：准备汇总分析
      setTables(prev => prev.map((t, i) => 
        i === index ? { 
          ...t, 
          analysisApiMessage: '批量处理完成，正在进行最终汇总...'
        } : t
      ));

      console.log(`所有批次处理完成，共${batchResults.length}个结果，开始汇总...`);

      // 汇总所有批次结果
      const summaryPrompt = `# 角色
你是一个信息汇总助手，负责将大模型每次的回复进行汇总，并按照特定格式输出全部内容。

## 技能
### 技能 1: 汇总回复信息
1. 仔细梳理大模型的回复内容，从中提取关键信息。
2. 按照以下格式进行汇总输出：
[一级标题]（XX人提及）
- 提出用户：
  * 姓名：xxx，系统号：xxx
  * 姓名：xxx，系统号：xxx
- 具体问题：
  1. 问题1（XX人提及）：
     - 问题描述：详细概括问题的核心内容
     - 代表性原文："xxx"（来自：姓名-系统号）
     - 代表性原文："xxx"（来自：姓名-系统号）
  2. 问题2（XX人提及）：
     - 问题描述：详细概括问题的核心内容
     - 代表性原文："xxx"（来自：姓名-系统号）
     - 代表性原文："xxx"（来自：姓名-系统号）

[负向反馈]
- 问题1（XX人提及）：
  * 问题描述：详细描述问题的具体情况
  * 提出用户：姓名-系统号
  * 原文内容："xxx"
- 问题2（XX人提及）：
  * 问题描述：详细描述问题的具体情况
  * 提出用户：姓名-系统号
  * 原文内容："xxx" 

## 限制:
- 输出内容必须严格按照给定格式组织，不得偏离框架要求。
- 确保汇总信息准确、清晰，能够真实反映大模型回复的关键要点。 

请汇总以下分析结果：
${batchResults.map((result, i) => `=== 批次${i + 1}分析结果 ===\n${result}`).join('\n\n')}`;

      const finalResponse = await fetch('https://ark.cn-beijing.volces.com/api/v3/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer af954c6d-7c97-4d29-926e-873807ed6032'
        },
        body: JSON.stringify({
          model: 'doubao-1-5-pro-256k-250115',
          messages: [
            { role: 'system', content: '你是一个信息汇总助手，负责将分析结果进行汇总整理。' },
            { role: 'user', content: summaryPrompt }
          ]
        })
      });

      if (!finalResponse.ok) {
        throw new Error(`最终汇总API请求失败: ${finalResponse.status} ${finalResponse.statusText}`);
      }

      const finalResult = await finalResponse.json();
      
      if (!finalResult.choices || !finalResult.choices[0] || !finalResult.choices[0].message) {
        throw new Error('最终汇总API返回数据格式错误');
      }

      const analysisResult = finalResult.choices[0].message.content;
      console.log('✓ 最终汇总完成，结果长度：', analysisResult.length);

      // 更新最终结果和成功状态
      setTables(prev => prev.map((table, i) => 
        i === index ? { 
          ...table, 
          isAnalyzing: false,
          analysis: analysisResult,
          analysisApiStatus: 'success',
          analysisApiMessage: '整体分析完成！'
        } : table
      ));

      message.success('整体分析完成！');

    } catch (error) {
      console.error('整体分析失败:', error);
      
      // 更新错误状态
      setTables(prev => prev.map((table, i) => 
        i === index ? { 
        ...table,
          isAnalyzing: false,
          analysisApiStatus: 'error',
          analysisApiMessage: `分析失败: ${error instanceof Error ? error.message : '未知错误'}`
        } : table
      ));
      
      message.error('整体分析失败，请重试');
    }
  };

  /**
   * @function processExcelFile
   * @description 处理Excel文件
   * @param {File} file - 上传的文件
   */
  const processExcelFile = async (file: File) => {
    setLoading(true);
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
          raw: false,  // 返回格式化的字符串
          defval: ''   // 空值默认为空字符串
        }) as ExcelData[];

        if (jsonData.length === 0) {
          message.error('Excel文件为空！');
          setLoading(false);
          return;
        }

        console.log('📊 原始数据示例：', jsonData.slice(0, 2));
        
        const allColumns = Object.keys(jsonData[0] as object);
        console.log('📑 所有列名：', allColumns);
        
        // 获取前三列（用户信息列）
        const firstThreeColumns = allColumns.slice(0, 3);
        console.log('👤 用户信息列：', firstThreeColumns);
        
        // 过滤剩余列
        const remainingColumns = allColumns.slice(3).filter(col => 
          !shouldSkipColumn(col) && hasEnoughUniqueValues(jsonData, col)
        );
        console.log('📝 待分析的列：', remainingColumns);

        const tables: ProcessedTable[] = [];
        
        for (const column of remainingColumns) {
          console.log(`\n开始处理列：${column}`);
          
          // 构建表格列定义
          const tableColumns = [
            ...firstThreeColumns.map(col => ({
              title: col,
              dataIndex: col,
              key: col,
              width: 150,
              ellipsis: true,
              fixed: col === firstThreeColumns[0] ? 'left' as const : undefined
            })),
            {
              title: column,
              dataIndex: column,
              key: column,
              width: 300
            }
          ];

          // 过滤并处理数据 - 按需求文档要求删除最后一列为空值或"无"的整行
          const tableData = jsonData
            .filter(row => {
              const value = row[column];
              // 按需求文档：删除最后一列为空值或"无"的整行
              return value !== undefined && value !== null && value !== '' && value !== '无';
            })
            .map((row: ExcelData, index: number) => {
              // 创建新的数据对象，确保数据对应关系正确
              const newRow: { [key: string]: string | number } = {
                key: index
              };

              // 添加用户信息列
              firstThreeColumns.forEach(col => {
                newRow[col] = row[col] || '';
              });

              // 添加当前分析的列
              newRow[column] = row[column];

              console.log(`处理第 ${index + 1} 行数据:`, {
                用户信息: firstThreeColumns.map(col => `${col}: ${newRow[col]}`),
                分析列: `${column}: ${newRow[column]}`
              });

              return newRow;
            });

          console.log(`✓ 列处理完成：${column}`);
          console.log(`- 总数据条数：${tableData.length}`);
          if (tableData.length > 0) {
            console.log('- 数据示例：');
            console.log(JSON.stringify(tableData[0], null, 2));
          }

          tables.push({
            name: `表格 - ${column}（共 ${tableData.length} 条数据）`,
            data: tableData,
            columns: tableColumns,
            isAnalyzing: false,
            commonTopics: [],
            isAnalyzingTopics: false
          });
        }

        setTables(tables);
        message.success(`Excel文件处理成功！共生成 ${tables.length} 个表格，请点击"分析"按钮分析具体问题。`);
      } catch (error) {
        console.error('处理Excel文件时发生错误：', error);
        message.error('处理Excel文件时发生错误！');
      } finally {
        setLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  /**
   * @function handleUpload
   * @description 处理文件上传
   * @param {UploadFile} file - 上传的文件信息
   */
  const handleUpload = (file: UploadFile) => {
    const fileType = file.name.split('.').pop()?.toLowerCase();
    if (fileType !== 'xlsx' && fileType !== 'xls') {
      message.error('请上传Excel文件！');
      return false;
    }
    processExcelFile(file as unknown as File);
    return false;
  };

  // 分析单条数据
  const analyzeTopicWithDeepSeek = async (content: string, topics: string[], columnName: string): Promise<string> => {
    try {
      // 如果内容为空，直接返回
      if (!content || content.trim() === '') {
        return '内容为空';
      }

      const prompt = `# 角色
你是一个数据分类分析专家，擅长判断数据内容所属类别，并能对特殊数据类型进行归纳总结。

## 技能
### 技能 1: 判断数据类别
1. 接收语义分析提供的数据内容。
2. 将数据内容与共性内容参考进行比对。
3. 如果数据内容属于共性内容参考中的某一类，输出这类的名称；若有多种符合，则输出多种。

### 技能 2: 归纳特殊类型
1. 若数据内容不属于共性内容参考中的任何一类。
2. 对数据内容进行总结归纳，得出它的特殊类型。
3. 若有多种特殊类型，则输出多种。

## 限制:
- 仅围绕语义分析提供的数据内容以及共性内容参考进行判断和归纳，不涉及其他无关话题。
- 输出内容应简洁明了，直接呈现判断结果或归纳出的特殊类型。

问题：${columnName}
答案内容：${content}
共性内容参考：${topics.join('、')}`;

      const response = await axios.post(
        ARK_API_URL,
        {
          model: MODEL_NAME,
          messages: [
            { 
              role: "system", 
              content: "你是一个数据分类分析专家，擅长判断数据内容所属类别，并能对特殊数据类型进行归纳总结。"
            },
            { role: "user", content: prompt }
          ]
        },
        {
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${ARK_API_KEY}`
          }
        }
      );

      if (response.data?.choices?.[0]?.message?.content) {
        return response.data.choices[0].message.content;
      }

      throw new Error('API返回数据格式错误');
    } catch (error: any) {
      console.error('分析单条数据失败:', error);
      if (error.response?.status === 401) {
        throw new Error('API认证失败，请检查API密钥');
      }
      if (error.response?.status === 429) {
        throw new Error('API请求过于频繁，请稍后重试');
      }
      throw new Error(`分析失败: ${error.message}`);
    }
  };

  /**
   * 处理共性内容分析
   * @param index 表格索引
   */
  const handleTopicsAnalysis = async (index: number) => {
    if (tables[index].isAnalyzingTopics) return;
    
    const topics = commonTopicsInput.split(/[、,，\n]/).filter(topic => topic.trim());
    if (topics.length === 0) {
      message.error('请输入至少一个共性内容');
      return;
    }

    // 保存当前表格的引用，避免在处理过程中状态变化导致的问题
    const currentTable = tables[index];

    // 更新状态：开始分析
    setTables(prev => prev.map((table, i) => 
      i === index ? { 
        ...table, 
        isAnalyzingTopics: true,
        commonTopics: topics,
        topicsApiStatus: 'analyzing',
        topicsApiMessage: '正在进行共性内容分析...'
      } : table
    ));

    try {
      const BATCH_SIZE = 300;
      const batches = [];
      
      // 创建批次
      for (let i = 0; i < currentTable.data.length; i += BATCH_SIZE) {
        batches.push(currentTable.data.slice(i, i + BATCH_SIZE));
      }

      // 更新API状态：开始批量处理
      setTables(prev => prev.map((t, i) => 
        i === index ? { 
          ...t, 
          topicsApiMessage: `开始批量处理，共${batches.length}批，每批最多${BATCH_SIZE}条数据`
        } : t
      ));

      // 从表格名称中提取实际的列名
      const columnName = currentTable.name.replace(/^表格 - /, '').replace(/（共 \d+ 条数据）$/, '');

      // 处理每个批次
      for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
        const batch = batches[batchIndex];
        
        // 更新进度
        setTables(prev => prev.map((t, i) => 
          i === index ? { 
            ...t, 
            topicsApiMessage: `正在处理第${batchIndex + 1}/${batches.length}批数据...`
          } : t
        ));

        // 处理批次中的每一行
        for (let rowIndex = 0; rowIndex < batch.length; rowIndex++) {
          const globalIndex = batchIndex * BATCH_SIZE + rowIndex;
          const row = batch[rowIndex];
          const content = row[columnName] || row[currentTable.name] || '';
          
          if (content) {
            try {
              const result = await analyzeTopicWithDeepSeek(content, topics, columnName);
              
              // 更新数据
              setTables(prev => prev.map((t, i) => {
                if (i === index) {
                  const newData = [...t.data];
                  newData[globalIndex] = { ...newData[globalIndex], 共性内容分析: result };
                  return { ...t, data: newData };
                }
                return t;
              }));

              // 每5行更新一次进度
              if ((rowIndex + 1) % 5 === 0) {
                setTables(prev => prev.map((t, i) => 
                  i === index ? { 
                    ...t, 
                    topicsApiMessage: `第${batchIndex + 1}批：已处理${rowIndex + 1}/${batch.length}行数据`
                  } : t
                ));
              }
            } catch (error) {
              console.error(`分析第${globalIndex + 1}行数据失败:`, error);
              // 标记失败但继续处理
              setTables(prev => prev.map((t, i) => {
                if (i === index) {
                  const newData = [...t.data];
                  newData[globalIndex] = { ...newData[globalIndex], 共性内容分析: '分析失败' };
                  return { ...t, data: newData };
                }
                return t;
              }));
            }
          }
        }
      }

      // 更新API状态：准备汇总分析
      setTables(prev => prev.map((t, i) => 
        i === index ? { 
          ...t, 
          topicsApiMessage: '数据处理完成，正在进行汇总分析...'
        } : t
      ));

      // 汇总分析 - 获取最新的表格数据
      const updatedTables = tables.map((t, i) => i === index ? { ...t, data: currentTable.data } : t);
      const updatedTable = updatedTables[index];
      
      const analysisData = updatedTable.data.map(row => ({
        用户信息: `${row['姓名'] || ''}-${row['系统号'] || ''}`,
        问题内容: row[columnName] || row[currentTable.name] || '',
        共性内容分析: row['共性内容分析'] || ''
      }));

      const summaryPrompt = `# 角色
你是一位资深的数据处理分析师，在数据清洗与预处理领域拥有深厚造诣，同时具备卓越的基于数据进行深入文本分析的能力。你以严谨、细致、高效的态度完成各项数据处理与分析任务。

## 技能
### 技能 1: 数据清洗与预处理
1. 接收输入数据后，运用专业方法全方位排查数据中存在的各类问题，涵盖缺失值、重复值、异常值等。深入剖析问题产生的原因，不局限于问题的发现。
2. 依据数据的特征、来源以及后续分析需求，精准挑选最合适的方法对这些问题进行妥善处理，保证数据质量达到高质量分析所需的严格标准。处理过程需详细记录文档，以便追溯。

### 技能 2: 文本分析与共性问题识别
1. 不管输入数据是否规范，针对表头的问题内容以及用户提供的回答展开深度文本分析。不仅进行词频统计、语义分析等常规操作，还运用关联分析等技术手段挖掘潜在文本关系。
2. 通过多种技术手段，精确识别其中的共性问题。共性问题提取要具体准确，避免过于宽泛的概括（例如不能简单概括为管理），对于难以准确归类的问题，单独列出说明。
3. 统计每个共性问题被提及的用户数量，详细罗列提出共性问题的所有用户姓名以及用户系统号，确保无遗漏。同时，记录每个用户提出问题的时间信息。
4. 针对共性问题下高频出现的内容进行系统总结，统计每个高频内容被提及的人数。除人数统计外，分析高频内容出现的趋势变化。
5. 从每个高频内容中选取有代表性的原文内容呈现出来，并备注输出该原文内容的用户姓名、系统号以及相关时间信息。

## 限制:
- 回答必须紧密围绕数据清洗、预处理以及文本分析相关任务，坚决拒绝回答无关话题。
- 输出内容要以清晰、符合逻辑的格式呈现共性问题、用户姓名、用户系统号、共性问题下高频出现内容的总结、全部被提及人数以及相关时间信息，不得有任何遗漏。 
- 千万不要遗漏提及共性问题的用户统计与输出，包括用户提出问题的时间信息。
注意：
- 一级标题包括：提及人数>=3共性内容 + 提及人数>=3的非共性内容
- 按提及人数从多到少排序，若人数相同，按照问题首次出现的时间先后排序
- 每个标题下必须包含用户清单和具体问题描述
- 负向内容必须全部呈现，不论提及人数多少
- 每个具体问题都要标注提及人数并引用代表性原文
- 不得因篇幅过长而省略任何问题或内容
- 所有满足条件的问题必须完整展示，包括所有用户信息和原文引用
- 确保输出所有相关内容，即使最终分析结果较长，要注重数据的完整性和准确性。 

## 输出格式：
[一级标题]（XX人提及）
- 提出用户：
  * 姓名：xxx，系统号：xxx
  * 姓名：xxx，系统号：xxx
- 具体问题：
  1. 问题1（XX人提及）：
     - 问题描述：xxx
     - 代表性原文："xxx"（来自：姓名-系统号）
     - 代表性原文："xxx"（来自：姓名-系统号）
  2. 问题2（XX人提及）：
     - 问题描述：xxx
     - 代表性原文："xxx"（来自：姓名-系统号）
     - 代表性原文："xxx"（来自：姓名-系统号）

[负向反馈]
- 问题1（XX人提及）：
  * 问题描述：xxx
  * 提出用户：姓名-系统号
  * 原文内容："xxx"
- 问题2（XX人提及）：
  * 问题描述：xxx
  * 提出用户：姓名-系统号
  * 原文内容："xxx" 

共性内容参考：${topics.join('、')}

请分析以下数据：
${JSON.stringify(analysisData, null, 2)}`;

      const response = await fetch('https://ark.cn-beijing.volces.com/api/v3/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer af954c6d-7c97-4d29-926e-873807ed6032'
        },
        body: JSON.stringify({
          model: 'doubao-1-5-pro-256k-250115',
          messages: [
            { role: 'system', content: '你是一个数据分析专家，擅长文本分析和共性问题识别。' },
            { role: 'user', content: summaryPrompt }
          ]
        })
      });

      if (!response.ok) {
        throw new Error(`API请求失败: ${response.status} ${response.statusText}`);
      }

      const result = await response.json();
      
      if (!result.choices || !result.choices[0] || !result.choices[0].message) {
        throw new Error('API返回数据格式错误');
      }

      const analysisResult = result.choices[0].message.content;

      // 更新最终结果和成功状态
      setTables(prev => prev.map((table, i) => 
        i === index ? { 
        ...table,
          isAnalyzingTopics: false,
          topicsAnalysis: analysisResult,
          topicsApiStatus: 'success',
          topicsApiMessage: '共性内容分析完成！'
        } : table
      ));

      message.success('共性内容分析完成！');
      setCommonTopicsInput('');

    } catch (error) {
      console.error('共性内容分析失败:', error);
      
      // 更新错误状态
      setTables(prev => prev.map((table, i) => 
        i === index ? { 
          ...table, 
          isAnalyzingTopics: false,
          topicsApiStatus: 'error',
          topicsApiMessage: `分析失败: ${error instanceof Error ? error.message : '未知错误'}`
        } : table
      ));
      
      message.error('共性内容分析失败，请重试');
    }
  };

  return (
    <div style={{ 
      padding: '32px'
    }}>
      <h1 style={{ 
        fontSize: '28px', 
        marginBottom: '32px',
        color: '#1f1f1f',
        fontWeight: '600',
        textAlign: 'center',
        position: 'relative'
      }}>问卷数据分析</h1>
      
      {tables.length === 0 && (
        <div style={{
          background: 'linear-gradient(135deg, #f0f7ff 0%, #e6f3ff 100%)',
          padding: '32px',
          borderRadius: '12px',
          marginBottom: '32px',
          border: '1px solid rgba(24, 144, 255, 0.1)',
          boxShadow: '0 2px 12px rgba(0, 0, 0, 0.04)'
        }}>
          <h2 style={{ 
            fontSize: '18px', 
            marginBottom: '24px', 
            color: '#1890ff',
            display: 'flex',
            alignItems: 'center',
            gap: '12px'
          }}>
            <span style={{ 
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: '32px',
              height: '32px',
              backgroundColor: '#1890ff',
              borderRadius: '50%',
              color: '#fff',
              fontSize: '16px',
              fontWeight: '500'
            }}>?</span>
            操作指引
          </h2>
          <ol style={{ 
            paddingLeft: '28px',
            margin: 0,
            fontSize: '15px',
            lineHeight: '2',
            color: '#262626'
          }}>
            <li>上传你的原始问卷excel表</li>
            <li>默认过滤选择题，找到你需要进行AI分析的问题</li>
            <li>对于这个问题有历史可参考的共性内容，可以上传后，在进行AI分析（会更准确）</li>
          </ol>
        </div>
      )}

      <Upload
        accept=".xlsx,.xls"
        beforeUpload={handleUpload}
        showUploadList={false}
        disabled={loading}
      >
        <Button 
          icon={<UploadOutlined />} 
          disabled={loading}
          type="primary"
          size="large"
          style={{
            height: '48px',
            padding: '0 32px',
            fontSize: '16px',
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
            margin: '0 auto'
          }}
        >
          {loading ? '处理中...' : '上传Excel文件'}
        </Button>
      </Upload>

      {loading && (
        <div style={{ 
          textAlign: 'center', 
          margin: '48px 0',
          padding: '48px',
          background: 'linear-gradient(135deg, #fafafa 0%, #f5f5f5 100%)',
          borderRadius: '12px'
        }}>
          <Spin tip="正在处理数据..." size="large" />
        </div>
      )}

      {tables.map((table, index) => (
        <div key={index} style={{ 
          marginTop: '32px',
          background: '#fff',
          padding: '24px',
          borderRadius: '12px',
          border: '1px solid #f0f0f0',
          boxShadow: '0 2px 12px rgba(0,0,0,0.04)'
        }}>
          <div style={{ 
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'space-between', 
            marginBottom: '24px',
            borderBottom: '1px solid #f0f0f0',
            paddingBottom: '20px'
          }}>
            <h3 style={{ 
              margin: 0, 
              color: '#262626',
              fontSize: '18px',
              fontWeight: '500'
            }}>{table.name}</h3>
            <div style={{ 
              display: 'flex', 
              alignItems: 'center',
              gap: '16px'
            }}>
            <div style={{ 
              display: 'flex', 
              alignItems: 'center', 
              gap: '16px'
            }}>
              <Input
                placeholder="输入一级共性内容，用、分隔"
                value={commonTopicsInput}
                onChange={(e) => setCommonTopicsInput(e.target.value)}
                style={{ 
                  width: '300px',
                  height: '40px',
                  borderRadius: '6px'
                }}
              />
              <Button
                type="primary"
                onClick={() => handleTopicsAnalysis(index)}
                loading={table.isAnalyzingTopics}
                disabled={table.isAnalyzingTopics || !commonTopicsInput.trim()}
                style={{ 
                  height: '40px',
                  borderRadius: '6px',
                  fontWeight: '500'
                }}
              >
                上传并分析
              </Button>
              </div>
              <Button 
                type="primary" 
                icon={<LineChartOutlined />} 
                onClick={() => handleAnalyze(index)}
                loading={table.isAnalyzing}
                disabled={table.isAnalyzing || !!table.analysis}
                style={{ 
                  height: '40px',
                  borderRadius: '6px',
                  fontWeight: '500'
                }}
              >
                {table.isAnalyzing ? '分析中...' : 
                 table.analysis ? '已分析' : '整体分析'}
              </Button>
            </div>
          </div>
          
          {/* API状态显示 */}
          {(table.topicsApiStatus || table.analysisApiStatus) && (
            <div style={{ marginBottom: '16px' }}>
              {/* 共性内容分析状态 */}
              {table.topicsApiStatus && (
                <div style={{
                  padding: '12px 16px',
                  borderRadius: '6px',
                  marginBottom: '8px',
                  backgroundColor: 
                    table.topicsApiStatus === 'success' ? '#f6ffed' :
                    table.topicsApiStatus === 'error' ? '#fff2f0' : '#f0f7ff',
                  border: `1px solid ${
                    table.topicsApiStatus === 'success' ? '#b7eb8f' :
                    table.topicsApiStatus === 'error' ? '#ffccc7' : '#91d5ff'
                  }`,
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px'
                }}>
                  <span style={{
                    color: 
                      table.topicsApiStatus === 'success' ? '#52c41a' :
                      table.topicsApiStatus === 'error' ? '#ff4d4f' : '#1890ff',
                    fontWeight: '500'
                  }}>
                    {table.topicsApiStatus === 'success' ? '✓' :
                     table.topicsApiStatus === 'error' ? '✗' : '⟳'}
                  </span>
                  <span style={{
                    color: 
                      table.topicsApiStatus === 'success' ? '#52c41a' :
                      table.topicsApiStatus === 'error' ? '#ff4d4f' : '#1890ff',
                    fontSize: '14px'
                  }}>
                    共性内容分析：{table.topicsApiMessage}
                  </span>
                </div>
              )}
              
              {/* 整体分析状态 */}
              {table.analysisApiStatus && (
                <div style={{
                  padding: '12px 16px',
                  borderRadius: '6px',
                  backgroundColor: 
                    table.analysisApiStatus === 'success' ? '#f6ffed' :
                    table.analysisApiStatus === 'error' ? '#fff2f0' : '#f0f7ff',
                  border: `1px solid ${
                    table.analysisApiStatus === 'success' ? '#b7eb8f' :
                    table.analysisApiStatus === 'error' ? '#ffccc7' : '#91d5ff'
                  }`,
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px'
                }}>
                  <span style={{
                    color: 
                      table.analysisApiStatus === 'success' ? '#52c41a' :
                      table.analysisApiStatus === 'error' ? '#ff4d4f' : '#1890ff',
                    fontWeight: '500'
                  }}>
                    {table.analysisApiStatus === 'success' ? '✓' :
                     table.analysisApiStatus === 'error' ? '✗' : '⟳'}
                  </span>
                  <span style={{
                    color: 
                      table.analysisApiStatus === 'success' ? '#52c41a' :
                      table.analysisApiStatus === 'error' ? '#ff4d4f' : '#1890ff',
                    fontSize: '14px'
                  }}>
                    整体分析：{table.analysisApiMessage}
                  </span>
                </div>
              )}
            </div>
          )}
          
          {/* 按需求文档要求：共性内容分析结果显示在问题和表格中间 */}
          {table.topicsAnalysis && (
            <div style={{ 
              backgroundColor: '#fafafa', 
              padding: '24px', 
              borderRadius: '12px',
              marginBottom: '24px',
              border: '1px solid #f0f0f0',
              boxShadow: 'inset 0 2px 8px rgba(0,0,0,0.02)'
            }}>
              <h4 style={{ 
                marginTop: 0, 
                color: '#262626',
                fontSize: '16px',
                fontWeight: '500',
                marginBottom: '16px',
                display: 'flex',
                alignItems: 'center',
                gap: '8px'
              }}>
                <LineChartOutlined style={{ color: '#1890ff' }} />
                共性内容分析结果
              </h4>
              <div style={{ 
                whiteSpace: 'pre-wrap', 
                wordWrap: 'break-word',
                margin: 0,
                fontFamily: 'inherit',
                fontSize: '14px',
                lineHeight: '1.8',
                maxHeight: '800px',
                overflowY: 'auto',
                backgroundColor: '#fff',
                padding: '20px',
                borderRadius: '8px',
                border: '1px solid #f0f0f0'
              }}>
                {table.topicsAnalysis.split('\n').map((line, i) => {
                  if (line.startsWith('[')) {
                    return <h3 key={i} style={{
                      color: '#1890ff',
                      fontSize: '16px',
                      marginTop: i === 0 ? 0 : '24px',
                      marginBottom: '16px',
                      fontWeight: '500'
                    }}>{line}</h3>;
                  } else if (line.startsWith('-')) {
                    return <div key={i} style={{
                      marginLeft: '20px',
                      marginBottom: '8px',
                      color: '#262626'
                    }}>{line}</div>;
                  } else if (line.startsWith('  *')) {
                    return <div key={i} style={{
                      marginLeft: '40px',
                      marginBottom: '4px',
                      color: '#595959'
                    }}>{line}</div>;
                  } else if (line.match(/^\s*\d+\./)) {
                    return <div key={i} style={{
                      marginLeft: '40px',
                      marginTop: '12px',
                      marginBottom: '8px',
                      color: '#262626',
                      fontWeight: '500'
                    }}>{line}</div>;
                  } else if (line.startsWith('     -')) {
                    return <div key={i} style={{
                      marginLeft: '60px',
                      marginBottom: '4px',
                      color: '#595959'
                    }}>{line}</div>;
                  } else {
                    return <div key={i} style={{
                      marginBottom: '4px',
                      color: '#262626'
                    }}>{line}</div>;
                  }
                })}
              </div>
            </div>
          )}

          {/* 按需求文档要求：整体分析结果显示在问题和表格中间 */}
          {table.analysis && (
            <div style={{ 
              backgroundColor: '#fafafa', 
              padding: '24px', 
              borderRadius: '12px',
              marginBottom: '24px',
              border: '1px solid #f0f0f0',
              boxShadow: 'inset 0 2px 8px rgba(0,0,0,0.02)'
            }}>
              <h4 style={{ 
                marginTop: 0, 
                color: '#262626',
                fontSize: '16px',
                fontWeight: '500',
                marginBottom: '16px',
                display: 'flex',
                alignItems: 'center',
                gap: '8px'
              }}>
                <LineChartOutlined style={{ color: '#1890ff' }} />
                整体分析结果
              </h4>
              <div style={{ 
                whiteSpace: 'pre-wrap', 
                wordWrap: 'break-word',
                margin: 0,
                fontFamily: 'inherit',
                fontSize: '14px',
                lineHeight: '1.8',
                maxHeight: '800px',
                overflowY: 'auto',
                backgroundColor: '#fff',
                padding: '20px',
                borderRadius: '8px',
                border: '1px solid #f0f0f0'
              }}>
                {table.analysis.split('\n').map((line, i) => {
                  if (line.startsWith('[')) {
                    return <h3 key={i} style={{
                      color: '#1890ff',
                      fontSize: '16px',
                      marginTop: i === 0 ? 0 : '24px',
                      marginBottom: '16px',
                      fontWeight: '500'
                    }}>{line}</h3>;
                  } else if (line.startsWith('-')) {
                    return <div key={i} style={{
                      marginLeft: '20px',
                      marginBottom: '8px',
                      color: '#262626'
                    }}>{line}</div>;
                  } else if (line.startsWith('  *')) {
                    return <div key={i} style={{
                      marginLeft: '40px',
                      marginBottom: '4px',
                      color: '#595959'
                    }}>{line}</div>;
                  } else if (line.match(/^\s*\d+\./)) {
                    return <div key={i} style={{
                      marginLeft: '40px',
                      marginTop: '12px',
                      marginBottom: '8px',
                      color: '#262626',
                      fontWeight: '500'
                    }}>{line}</div>;
                  } else if (line.startsWith('     -')) {
                    return <div key={i} style={{
                      marginLeft: '60px',
                      marginBottom: '4px',
                      color: '#595959'
                    }}>{line}</div>;
                  } else {
                    return <div key={i} style={{
                      marginBottom: '4px',
                      color: '#262626'
                    }}>{line}</div>;
                  }
                })}
              </div>
            </div>
          )}

          {/* 数据表格 */}
          {table.commonTopics.length === 0 && (
            <Table
              columns={table.columns}
              dataSource={table.data}
              scroll={{ x: true }}
              pagination={{ 
                pageSize: 10,
                showTotal: (total) => `共 ${total} 条数据`,
                showSizeChanger: true,
                showQuickJumper: true
              }}
              size="middle"
              bordered
              style={{
                borderRadius: '8px',
                overflow: 'hidden',
                marginBottom: '24px'
              }}
            />
          )}
          
          {/* 共性内容分析的数据表格 */}
          {table.commonTopics.length > 0 && (
            <div style={{ 
              backgroundColor: '#fafafa', 
              padding: '24px', 
              borderRadius: '12px',
              marginBottom: '24px',
              border: '1px solid #f0f0f0',
              boxShadow: 'inset 0 2px 8px rgba(0,0,0,0.02)'
            }}>
              <h4 style={{ 
                marginTop: 0, 
                color: '#262626',
                fontSize: '16px',
                fontWeight: '500',
                marginBottom: '16px',
                display: 'flex',
                alignItems: 'center',
                gap: '8px'
              }}>
                <LineChartOutlined style={{ color: '#1890ff' }} />
                共性内容分析数据
              </h4>
              <div style={{
                marginBottom: '16px',
                padding: '12px',
                backgroundColor: '#fff',
                borderRadius: '6px',
                border: '1px solid #f0f0f0'
              }}>
                <div style={{
                  color: '#595959',
                  marginBottom: '8px'
                }}>已设置的共性内容：</div>
                <div style={{
                  display: 'flex',
                  flexWrap: 'wrap',
                  gap: '8px'
                }}>
                  {table.commonTopics.map((topic, i) => (
                    <span key={i} style={{
                      padding: '4px 12px',
                      backgroundColor: '#f0f7ff',
                      borderRadius: '4px',
                      color: '#1890ff',
                      fontSize: '14px'
                    }}>{topic}</span>
                  ))}
                </div>
              </div>
          <Table
            columns={table.columns}
            dataSource={table.data}
            scroll={{ x: true }}
            pagination={{ 
              pageSize: 10,
              showTotal: (total) => `共 ${total} 条数据`,
              showSizeChanger: true,
              showQuickJumper: true
            }}
            size="middle"
            bordered
            style={{
              borderRadius: '8px',
              overflow: 'hidden'
            }}
          />
            </div>
          )}
        </div>
      ))}
    </div>
  );
};

export default ExcelProcessor; 