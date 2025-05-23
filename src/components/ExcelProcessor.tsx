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
}

/**
 * @interface ExcelData
 * @description Excel数据接口
 */
interface ExcelData {
  [key: string]: string | number;
}

// DeepSeek API配置
const DEEPSEEK_API_KEY = process.env.REACT_APP_DEEPSEEK_API_KEY || '';
const DEEPSEEK_API_URL = process.env.REACT_APP_DEEPSEEK_API_URL || 'https://api.deepseek.com/v1/chat/completions';

/**
 * @component ExcelProcessor
 * @description 问卷数据分析组件
 */
const ExcelProcessor: React.FC = () => {
  const [processedTables, setProcessedTables] = useState<ProcessedTable[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [apiStatus, setApiStatus] = useState<{ message: string; type: 'info' | 'success' | 'error' | ''; timestamp: number }>({
    message: '',
    type: '',
    timestamp: 0
  });
  const [commonTopicsInput, setCommonTopicsInput] = useState<string>('');

  /**
   * @function shouldSkipColumn
   * @description 判断是否应该跳过某列
   * @param {string} columnName - 列名
   * @returns {boolean}
   */
  const shouldSkipColumn = (columnName: string): boolean => {
    const skipPatterns = [
      /单选/,
      /多选/,
      /开始时间/,
      /结束时间/,
      /组织编码/,
      /组织信息/,
      /ucid/i,
      /岗位名称/,
      /公司所在城市/,
      /工作所在城市/,
      /品牌/,
      /门店信息/,
      /条线/,
      /所属组织/,
    ];

    if (/[单多]选/.test(columnName)) {
      const otherText = columnName.replace(/[单多]选/, '').trim();
      if (!otherText) {
        return true;
      }
    }

    return skipPatterns.some(pattern => pattern.test(columnName));
  };

  /**
   * @function hasEnoughUniqueValues
   * @description 检查列是否有足够的非空唯一值
   * @param {ExcelData[]} data - Excel数据
   * @param {string} columnName - 列名
   * @returns {boolean}
   */
  const hasEnoughUniqueValues = (data: ExcelData[], columnName: string): boolean => {
    const nonEmptyValues = data
      .map(row => row[columnName])
      .filter(value => value !== undefined && value !== null && value !== '');
    
    const uniqueValues = Array.from(new Set(nonEmptyValues));
    return uniqueValues.length > 10;
  };

  /**
   * @function analyzeDataWithDeepSeek
   * @description 使用DeepSeek API分析数据
   * @param {any[]} data - 表格数据
   * @param {string} columnName - 列名
   * @returns {Promise<string>}
   */
  const analyzeDataWithDeepSeek = async (data: any[], columnName: string): Promise<string> => {
    try {
      setApiStatus({ message: `开始分析列: ${columnName}`, type: 'info', timestamp: Date.now() });
      console.log('\n========================================');
      console.log(`📊 开始分析列: ${columnName}`);
      console.log(`📝 数据条数: ${data.length}`);
      console.log('----------------------------------------');
      console.log('🚀 正在调用 DeepSeek API...');

      const prompt = `作为资深数据处理分析师，请对以下数据进行深入分析。列名为"${columnName}"，数据内容如下：
${JSON.stringify(data, null, 2)}

分析要求：

1. 数据清洗与预处理
   - 全面排查数据中的缺失值、重复值、异常值等问题
   - 说明发现的数据质量问题及处理方法
   - 提供清洗后的有效数据量

2. 文本分析与共性问题识别
   - 基于问题"${columnName}"的内容，对用户回答进行深度文本分析
   - 通过词频统计和语义分析，识别具体的共性问题（避免过于宽泛的概括）
   - 对每个共性问题进行详细分析：
     a) 统计提出该问题的用户总数
     b) 列出所有提出该问题的用户信息（姓名和系统号）
     c) 总结该问题下的高频内容，并统计每个高频内容的提及人数
     d) 选取代表性的原文内容，并注明提出者的姓名和系统号

输出格式要求：
1. 数据质量报告
   - 原始数据量：
   - 发现的问题：
   - 处理方法：
   - 有效数据量：

2. 共性问题分析（按照提及人数从多到少排序）
   [共性问题1]
   - 提及总人数：XX人
   - 提出用户清单：
     * 姓名：xxx，系统号：xxx
     * 姓名：xxx，系统号：xxx
   - 高频内容分析：
     a) 内容主题1（XX人提及）：
        - 代表性原文："xxx"（来自：姓名-系统号）
     b) 内容主题2（XX人提及）：
        - 代表性原文："xxx"（来自：姓名-系统号）

   [共性问题2]
   ...（按相同格式继续）

注意事项：
- 确保每个共性问题下都完整列出所有提出该问题的用户信息
- 高频内容必须具体明确，避免笼统表述
- 选取的代表性原文要能准确反映问题特点`;

      const startTime = Date.now();
      setApiStatus({ message: '正在调用 API，请稍候...', type: 'info', timestamp: Date.now() });
      
      const requestData = {
        model: "deepseek-reasoner",
        messages: [
          { 
            role: "system", 
            content: `# 角色
你是一位资深的数据处理分析师，在数据清洗与预处理领域经验丰富，同时具备强大的基于数据进行深入文本分析的能力。

## 技能
### 技能 1: 数据清洗与预处理
1. 接收输入数据后，全面细致地排查数据中存在的各类问题，包括但不限于缺失值、重复值、异常值等。
2. 依据数据特点，运用最合适的方法对这些问题进行处理，务必使数据质量达到高质量分析的严格要求。

### 技能 2: 文本分析与共性问题识别
1. 无论输入数据是否规范，基于表头的问题内容以及用户提供的回答展开深度文本分析。
2. 通过词频统计、语义分析等先进技术手段，精准识别其中的共性问题。共性问题提取要具体准确，避免过于宽泛的概括（例如不能简单概括为管理）。
3. 统计每个共性问题被提及的用户数量，并详细列出提出共性问题的所有用户姓名以及用户系统号，确保无遗漏。
4. 针对共性问题下高频出现的内容进行系统总结，统计每个高频内容被提及的人数。
5. 从每个高频内容中选取有代表性的原文内容呈现出来，并备注输出该原文内容的用户姓名和系统号。

## 限制:
- 回答必须紧密围绕数据清洗、预处理以及文本分析相关任务，坚决拒绝回答无关话题。
- 输出内容要以清晰、符合逻辑的格式呈现共性问题、用户姓名、用户系统号、共性问题下高频出现内容的总结以及全部被提及人数，不得有任何遗漏。 
- 千万不要有遗漏提及共性问题的用户统计与输出`
          },
          { role: "user", content: prompt }
        ],
        temperature: 0.3
      };

      console.log('📡 API请求配置：', {
        url: DEEPSEEK_API_URL,
        model: requestData.model,
        temperature: requestData.temperature
      });

      const response = await axios.post(
        DEEPSEEK_API_URL,
        requestData,
        {
          headers: {
            'Authorization': `Bearer ${DEEPSEEK_API_KEY}`,
            'Content-Type': 'application/json',
          }
        }
      );

      const endTime = Date.now();
      const duration = ((endTime - startTime) / 1000).toFixed(2);

      setApiStatus({ message: `分析完成！用时：${duration}秒`, type: 'success', timestamp: Date.now() });
      console.log('✅ API调用成功！');
      console.log(`⏱️ 用时: ${duration}秒`);
      
      if (response.data?.choices?.[0]?.message?.content) {
        const result = response.data.choices[0].message.content;
        return result;
      }

      setApiStatus({ message: '未获取到有效的分析结果', type: 'error', timestamp: Date.now() });
      return '无分析结果';
    } catch (error) {
      console.error('\n❌ DeepSeek API调用失败:', error);
      const errorMessage = axios.isAxiosError(error) 
        ? error.response?.data?.error?.message || error.message
        : '未知错误';
      setApiStatus({ message: `API调用失败: ${errorMessage}`, type: 'error', timestamp: Date.now() });
      return '数据分析失败，请稍后重试';
    }
  };

  /**
   * @function handleAnalyze
   * @description 处理单个表格的分析请求
   * @param {number} index - 表格索引
   */
  const handleAnalyze = async (index: number) => {
    const tables = [...processedTables];
    const table = tables[index];
    
    // 如果已经分析过或正在分析中，则不再执行
    if (table.isAnalyzing || table.analysis) {
      message.info('该数据已经分析过了');
      return;
    }

    // 设置分析中状态
    table.isAnalyzing = true;
    setProcessedTables(tables);

    try {
      const columnName = table.columns[table.columns.length - 1].title;
      console.log('\n========================================');
      console.log(`开始分析表格 ${index + 1}/${tables.length}`);
      console.log(`列名: ${columnName}`);
      console.log(`数据量: ${table.data.length} 条`);
      console.log('----------------------------------------');

      const analysis = await analyzeDataWithDeepSeek(table.data, columnName);
      
      console.log('分析完成！');
      console.log('========================================\n');

      // 更新分析结果
      tables[index] = {
        ...table,
        analysis,
        isAnalyzing: false
      };
      setProcessedTables(tables);
    } catch (error) {
      console.error('分析过程出错:', error);
      message.error('分析失败，请重试');
      // 重置分析状态
      tables[index] = {
        ...table,
        isAnalyzing: false
      };
      setProcessedTables(tables);
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

          // 过滤并处理数据
          const tableData = jsonData
            .filter(row => {
              const value = row[column];
              return value !== undefined && value !== null && value !== '';
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

        setProcessedTables(tables);
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
      const prompt = `作为资深数据分析师，请分析以下内容属于哪些一级共性内容，如果不属于任何已知共性内容，请归纳总结。

问题：${columnName}
答案内容：${content}
已知一级共性内容：${topics.join('、')}

请按以下格式输出：
1. 如果属于一级共性内容：直接列出所属的一级共性内容，多个用、分隔
2. 如果不属于任何已知共性内容：总结其核心内容

注意：
- 输出必须简洁，不要有多余的解释
- 如果属于多个共性内容，请全部列出
- 如果不属于任何共性内容，总结时要简明扼要`;

      const response = await axios.post(
        DEEPSEEK_API_URL,
        {
          model: "deepseek-reasoner",
          messages: [
            { 
              role: "system", 
              content: "你是一位专注于文本分类的数据分析师，擅长准确判断内容的归属类别。请只输出分类结果，不要有任何多余的解释。"
            },
            { role: "user", content: prompt }
          ],
          temperature: 0.3
        },
        {
          headers: {
            'Authorization': `Bearer ${DEEPSEEK_API_KEY}`,
            'Content-Type': 'application/json',
          }
        }
      );

      return response.data?.choices?.[0]?.message?.content || '分析失败';
    } catch (error) {
      console.error('分析单条数据失败:', error);
      return '分析失败';
    }
  };

  // 处理共性内容分析
  const handleTopicsAnalysis = async (index: number) => {
    const tables = [...processedTables];
    const table = tables[index];
    
    if (table.isAnalyzingTopics) {
      return;
    }

    const topics = commonTopicsInput.split('、').filter(t => t.trim());
    if (topics.length === 0) {
      message.error('请输入共性内容，用、分隔');
      return;
    }

    table.isAnalyzingTopics = true;
    table.commonTopics = topics;
    setProcessedTables(tables);

    try {
      setApiStatus({ message: '正在分析每条数据...', type: 'info', timestamp: Date.now() });

      // 添加分析结果列
      const newColumns = [...table.columns];
      newColumns.push({
        title: '共性内容分析',
        dataIndex: 'topicAnalysis',
        key: 'topicAnalysis',
        width: 200,
        ellipsis: true
      });

      // 分析每条数据
      const columnName = table.columns[table.columns.length - 1].title;
      const analyzedData = await Promise.all(
        table.data.map(async (row) => {
          const content = row[columnName];
          const analysis = await analyzeTopicWithDeepSeek(content, topics, columnName);
          return {
            ...row,
            topicAnalysis: analysis
          };
        })
      );

      // 更新表格数据
      tables[index] = {
        ...table,
        data: analyzedData,
        columns: newColumns,
        isAnalyzingTopics: false
      };
      setProcessedTables(tables);

      // 分析完整表格
      handleAnalyze(index);

      setApiStatus({ message: '共性内容分析完成！', type: 'success', timestamp: Date.now() });
    } catch (error) {
      console.error('共性内容分析失败:', error);
      message.error('分析失败，请重试');
      tables[index].isAnalyzingTopics = false;
      setProcessedTables(tables);
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
      
      {processedTables.length === 0 && (
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

      {apiStatus.message && (
        <div style={{ 
          padding: '16px 24px',
          marginTop: '24px',
          backgroundColor: apiStatus.type === 'success' ? '#f6ffed' : 
                         apiStatus.type === 'error' ? '#fff2f0' : 
                         '#e6f7ff',
          border: `1px solid ${
            apiStatus.type === 'success' ? '#b7eb8f' : 
            apiStatus.type === 'error' ? '#ffccc7' : 
            '#91d5ff'
          }`,
          borderRadius: '8px',
          marginBottom: '24px',
          boxShadow: '0 2px 8px rgba(0,0,0,0.04)'
        }}>
          <div style={{ 
            display: 'flex', 
            alignItems: 'center',
            gap: '8px',
            color: apiStatus.type === 'success' ? '#52c41a' : 
                   apiStatus.type === 'error' ? '#ff4d4f' : 
                   '#1890ff',
            fontSize: '14px'
          }}>
            {apiStatus.type === 'success' ? '✅' : 
             apiStatus.type === 'error' ? '❌' : 
             '🔄'} {apiStatus.message}
          </div>
        </div>
      )}

      {processedTables.map((table, index) => (
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
              <Button 
                type="primary" 
                icon={<LineChartOutlined />} 
                onClick={() => handleAnalyze(index)}
                loading={table.isAnalyzing}
                disabled={table.isAnalyzing || !!table.analysis || !table.commonTopics}
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
                数据分析结果
              </h4>
              <pre style={{ 
                whiteSpace: 'pre-wrap', 
                wordWrap: 'break-word',
                margin: 0,
                fontFamily: 'inherit',
                fontSize: '14px',
                lineHeight: '1.8',
                maxHeight: '500px',
                overflowY: 'auto',
                backgroundColor: '#fff',
                padding: '20px',
                borderRadius: '8px',
                border: '1px solid #f0f0f0'
              }}>
                {table.analysis}
              </pre>
            </div>
          )}

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
      ))}
    </div>
  );
};

export default ExcelProcessor; 