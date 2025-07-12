import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import { bitable, IFieldMeta, ITable, ITableMeta } from "@lark-base-open/js-sdk";
import {
  Button,
  Card,
  Table,
  Tag,
  Toast,
  Modal,
  Progress,
  Upload,
  Typography,
  Form
} from '@douyinfe/semi-ui';
import * as XLSX from 'xlsx';
import { BaseFormApi } from '@douyinfe/semi-foundation/lib/es/form/interface';

const { Title, Text } = Typography;

const ExcelImportPlugin = () => {
  const [excelData, setExcelData] = useState<any[]>([]);
  const [fileName, setFileName] = useState<string>('');

  const [tableMetaList, setTableMetaList] = useState<ITableMeta[]>();
  const [tableId, setTableId] = useState<string>();
  const [table, setTable] = useState<ITable>();
  const [tableFields, setTableFields] = useState([]);


  // const [mapping, setMapping] = useState({});
  const [comparisonResult, setComparisonResult] = useState<{
    toAdd: any[]
    toDelete: any[]
  } | null>(null);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const fileInputRef = useRef<any>(null);
  const formApi = useRef<BaseFormApi>();

  useEffect(() => {
    Promise.all([bitable.base.getTableMetaList(), bitable.base.getSelection()])
      .then(([metaList, selection]) => {
        setTableMetaList(metaList);
        formApi.current?.setValues({ table: selection.tableId });
      });
  }, []);

  // 获取表格实例
  const getTable = useCallback(
    async () => {
      if (tableId) {
        const table = await bitable.base.getTableById(tableId);
        if (table) {
          setTable(table);
        }
      }
    },
    [tableId]
  )

  useEffect(() => {
    getTable();
  }, [getTable])


  // 插入数据格式转换
  const convertDataForInsertion = useCallback(
    (
      records: {
        fields: { [key: string]: string }
      }[],
      fieldMetaList: IFieldMeta[]
    ) => {
      // 创建字段元数据映射
      const fieldMetaMap = new Map(fieldMetaList.map(meta => [meta.id, meta]));

      return records.map(record => {
        const convertedFields: { [key: string]: any } = {};

        for (const [fieldId, value] of Object.entries(record.fields)) {
          const meta = fieldMetaMap.get(fieldId);
          if (!meta) continue; // 跳过不存在的字段

          // 处理空值
          if (value === null || value === undefined || value === "") {
            convertedFields[fieldId] = null;
            continue;
          }

          // 根据字段类型转换数据
          switch (meta.type) {
            // 数字字段
            case 2:
              convertedFields[fieldId] = Number(value) || 0;
              break;

            // 单选字段
            case 3:
              // 确保值在选项中存在
              const option = (meta?.property as any)?.options?.find((opt:any) => opt.name === value);
              convertedFields[fieldId] = option || null;
              break;

            // 多选字段
            case 4:
              // 多选字段需要数组格式
              const lawyer = value?.split('、');
              const arr: any[] = [];
              if (lawyer.length) {
                lawyer.forEach((lw) => {
                  const option1 = (meta?.property as any)?.options?.find((opt: any) => opt.name === lw);
                  if (option1) {
                    arr.push(option1)
                  }
                })
              }
              convertedFields[fieldId] = arr;
              break;

            // 日期字段
            case 5:
              // 尝试转换日期格式
              try {
                // 处理不同格式的日期
                let dateStr = value;
                if (typeof value === 'string') {
                  // 统一替换分隔符
                  dateStr = value.replaceAll('-', '/');
                }
                convertedFields[fieldId] = Date.parse(dateStr);
              } catch (e) {
                convertedFields[fieldId] = null;
              }
              break;

            // 其他类型保持原样
            default:
              convertedFields[fieldId] = value;
          }
        }

        return { fields: convertedFields };
      });
    },
    []
  )

  // 插入数据
  const safeAddRecords = useCallback(
    async (
      records: {
        fields: { [key: string]: string }
      }[],
      fieldMeta: IFieldMeta[]
    ) => {
      const processedRecords = convertDataForInsertion(records, fieldMeta);
      // 2. 分批处理（一次插入上限为200）
      const batchSize = 200;
      const results = [];
      if (!table) {
        Toast.error("获取表格实例失败，请稍后重试");
        return
      }
      for (let i = 0; i < processedRecords.length; i += batchSize) {
        const batch = processedRecords.slice(i, i + batchSize);
        try {
          const result = await table.addRecords(batch);
          results.push(...result);
        } catch (batchError) {
          console.error(`批处理错误 (记录 ${i}-${i + batchSize}):`, batchError);
          // 3. 单条重试（定位具体错误记录）
          for (const singleRecord of batch) {
            try {
              const singleResult = await table.addRecord(singleRecord);
              results.push(singleResult);
            } catch (singleError) {
              console.error('单条记录插入失败:', {
                record: singleRecord,
                error: singleError
              });
            }
          }
        }
      }

      return results;
    },
    [table]
  )

  // 删除数据
  const safeDeleteRecords = useCallback(
    async (
      recordIds: string[]
    ) => {
      const batchSize = 200;
      const results = [];
      if (!table) {
        Toast.error("获取表格实例失败，请稍后重试");
        return
      }
      for (let i = 0; i < recordIds.length; i += batchSize) {
        const batch = recordIds.slice(i, i + batchSize);
        try {
          const result = await table.deleteRecords(batch);
        } catch (batchError) {
          console.error(`批量删除处理错误 (记录 ${i}-${i + batchSize}):`, batchError);
          // 3. 单条重试（定位具体错误记录）
          for (const singleRecord of batch) {
            try {
              const singleResult = await table.deleteRecord(singleRecord);
              results.push(singleResult);
            } catch (singleError) {
              console.error('单条记录删除失败:', {
                record: singleRecord,
                error: singleError
              });
            }
          }
        }
      }
    },
    [table]
  )


  // 同步记录
  const syncRecords = useCallback(
    async (params: {
      action: 'add' | 'delete'
      data: any,
      progressCallback?: () => void
    }) => {
      const { action, data } = params;
      // return
      if (!table) {
        Toast.error('获取表格实例出错');
        return
      };
      // try {
      let processed = 0;
      const total = data.length;

      switch (action) {
        case 'add':
          // 批量添加记录
          const fieldIdList = await table.getFieldMetaList();
          const idMap: { [key: string]: string } = {};
          const keys = Object.keys(data?.[0] || {});
          keys.forEach((key) => {
            idMap[key] = fieldIdList?.find(it => it?.name === key)?.id || '';
          })
          const finalFields = data.map((item: { [key: string]: any }) => {
            const obj: { [key: string]: any } = {};
            Object.keys(item).forEach(key => {
              const id = idMap[key];
              if (id) {
                obj[id] = item[key];
              }
            })
            return {
              fields: obj
            }
          })
          console.log("批量增加-finalFields", finalFields);
          console.log("批量增加-idMap", idMap);
          await safeAddRecords(finalFields, fieldIdList)
          processed = total;
          break;

        case 'delete':
          // 批量删除记录
          const finalRecord: string[] = data?.map((it: any) => it?.recordId);
          console.log("批量删除", {
            data,
            finalRecord
          });
          await safeDeleteRecords(finalRecord);
          processed = total;
          break;

        default:
          break;
      }
    },
    [table]
  )

  // 读取Excel文件
  const readExcelFile = (file: any) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e: any) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet);
          resolve(jsonData);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  };


  // 选择并读取Excel文件
  const handleFileChange = async (e: any) => {
    const file = e.target.files[0];
    if (!file) return;
    console.log("file", file)
    setLoading(true);
    Toast.info('正在解析Excel文件...');

    try {
      const data: any[] = await readExcelFile(file) as any;
      setExcelData(data);
      setFileName(file?.name);
      Toast.success(`成功解析Excel文件，共${data.length}条数据`);

      // 初始化字段映射（自动匹配名称相同的字段）
      const initialMapping = {};
      if (data.length > 0) {
        Object.keys(data[0]).forEach(excelField => {
        });
      }
      // setMapping(initialMapping);
    } catch (error: any) {
      Toast.error(`文件解析失败: ${error.message}`);
    } finally {
      setLoading(false);
      e.target.value = ''; // 重置文件输入
    }
  };


  const getTableRecords = useCallback(
    async (prePageToken?: string, preRecords?: any[]) => {
      if (tableId) {
        try {
          const table = await bitable.base.getTableById(tableId);
          // const fieldMetaList = await table.getFieldMetaList();
          const params: any = { pageSize: 200 };
          if (prePageToken) {
            params.pageToken = prePageToken;
          }
          const { hasMore, pageToken, records } = await table.getRecords(params);
          console.log("多维表格数据获取结果", {
            hasMore,
            pageToken,
            records
          });
          const tempData = [...(preRecords || []), ...records]
          if (!hasMore) {
            return tempData
          }
          const data: any[] = await getTableRecords(pageToken, tempData)
          return data;
        } catch (error) {
          console.log("获取多维表格出错", error)
          return []
        }
      }
      return []
    },
    [tableId]
  )

  const getfileId = useCallback(async () => {
    if (tableId) {
      const table = await bitable.base.getTableById(tableId);
      const fieldMetaList = await table.getFieldMetaList();
      const id = fieldMetaList?.find(it => it?.name === '客户号')?.id;
      return id;
    }
    return ''
  }, [tableId])


  // 对比Excel和多维表格数据
  const compareData = async () => {
    if (excelData.length === 0) {
      Toast.warning('请先上传Excel文件');
      return;
    }
    if (!tableId) {
      Toast.warning('请先选择多维表格');
      return;
    }

    setLoading(true);
    Toast.info('正在对比数据...');

    try {
      // 获取多维表格数据
      const tableRecords: any = await getTableRecords();
      console.log("数据对比-tableRecords", tableRecords)
      const fileId = await getfileId();
      if (!fileId) {
        Toast.error('获取合同号字段id出错');
        return;
      }
      const idField = '客户号';
      const excelMap = new Map(excelData.map(item => [item[idField], item]));
      const newTableArr: any[] = [];
      tableRecords?.forEach((item: any) => {
        const key = item?.fields?.[fileId]?.[0]?.text;
        if (key) {
          newTableArr.push([
            key, item
          ])
        }
      })
      const tableMap = new Map(newTableArr);

      console.log("数据对比-excelMap", excelMap)
      console.log("数据对比-newTableMap", tableMap)
      const toAdd: any[] = [];
      const toDelete: any[] = [];
      // 识别需要新增的数据
      excelMap.forEach((excelItem, id) => {
        if (!tableMap.has(id)) {
          toAdd.push(excelItem);
        }
      });
      // 识别需要删除的数据
      tableMap.forEach((record, id) => {
        if (!excelMap.has(id)) {
          toDelete.push(record);
        }
      });

      console.log("对比结果", {
        toAdd,
        toDelete
      })

      setComparisonResult({
        toAdd,
        toDelete,
      });

      Toast.success('数据对比完成！');
    } catch (error: any) {
      console.log("数据对比失败error", error)
      Toast.error(`数据对比失败: ${error.message}`);
    } finally {
      setLoading(false);
    }
  };

  // 执行数据同步
  const executeSync = async (deleteMissing = false) => {
    if (!comparisonResult) {
      Toast.warning('请先进行数据对比');
      return;
    }

    setLoading(true);
    setProgress(0);

    try {
      const { toAdd, toDelete } = comparisonResult;
      let totalSteps = toAdd.length;
      if (deleteMissing) {
        totalSteps += toDelete.length;
      }

      // 添加新记录
      if (toAdd.length > 0) {
        Toast.info(`正在添加${toAdd.length}条新记录...`);
        await syncRecords({
          action: 'add',
          data: toAdd,
        })
      }

      // 删除多余记录
      if (deleteMissing && toDelete.length > 0) {
        Toast.info(`正在删除${toDelete.length}条多余记录...`);
        await syncRecords({
          action: 'delete',
          data: toDelete
        })
      }

      Toast.success(`操作完成！新增:${toAdd.length}, 删除:${deleteMissing ? toDelete.length : 0}`);
    } catch (error:any) {
      Toast.error(`同步失败: ${error.message}`);
    } finally {
      setLoading(false);
      setProgress(0);
    }
  };

  // 渲染字段映射UI
  // const renderFieldMapping = () => {
  //   if (excelData.length === 0 || tableFields.length === 0) return null;

  //   const excelHeaders = Object.keys(excelData[0]);

  //   return (
  //     <Card title="字段映射" style={{ marginTop: 20 }}>
  //       <Table
  //         dataSource={excelHeaders.map(header => ({ header }))}
  //         columns={[
  //           { title: 'Excel字段', dataIndex: 'header' },
  //           {
  //             title: '映射到多维表格字段',
  //             render: (_, { header }) => (
  //               <select
  //                 value={mapping[header] || ''}
  //                 onChange={(e) => setMapping({ ...mapping, [header]: e.target.value })}
  //                 style={{ width: '100%', padding: '8px' }}
  //               >
  //                 <option value="">-- 请选择 --</option>
  //                 {tableFields.map(field => (
  //                   <option key={field.id} value={field.id}>{field.name}</option>
  //                 ))}
  //               </select>
  //             )
  //           }
  //         ]}
  //         pagination={false}
  //       />
  //     </Card>
  //   );
  // };

  // 渲染对比结果
  const renderComparisonResult = useMemo(() => {
    if (!comparisonResult) return null;

    const { toAdd, toDelete } = comparisonResult;

    return (
      <Card title="对比结果" style={{ marginTop: 20 }}>
        <div style={{ display: 'flex', justifyContent: 'space-around', marginBottom: 20 }}>
          <div style={{ textAlign: 'center' }}>
            <Tag color="green" size="large">新增</Tag>
            <Title heading={3}>{toAdd.length}</Title>
          </div>
          {/* <div style={{ textAlign: 'center' }}>
            <Tag color="blue" size="large">更新</Tag>
            <Title heading={3}>{toUpdate.length}</Title>
          </div> */}
          <div style={{ textAlign: 'center' }}>
            <Tag color="red" size="large">删除</Tag>
            <Title heading={3}>{toDelete.length}</Title>
          </div>
          {/* <div style={{ textAlign: 'center' }}>
            <Tag color="grey" size="large">总记录</Tag>
            <Title heading={3}>{totalRecords}</Title>
          </div> */}
        </div>

        <div style={{ display: 'flex', gap: 10 }}>
          <Button
            theme="solid"
            type="primary"
            onClick={() => executeSync(false)}
            loading={loading}
          >
            合并数据（不删除）
          </Button>
          <Button
            theme="solid"
            type="danger"
            onClick={() => executeSync(true)}
            loading={loading}
          >
            合并并删除多余数据
          </Button>
        </div>
      </Card>
    );
  }, [comparisonResult])


  useEffect(() => {
    if (excelData?.length) {
      console.log("fileInputRef.current", fileInputRef.current)
    }
  }, [excelData?.length])

  return (
    <div style={{ maxWidth: 1000, margin: '0 auto', padding: 20 }}>
      <Title heading={2}>Excel数据导入与多维表格对比</Title>
      <Text type="secondary">将Excel数据与多维表格数据进行对比、合并和导入</Text>

      <Card
        title="第一步：上传Excel文件"
        style={{ marginTop: 20 }}
        headerExtraContent={
          <Button
            onClick={() => fileInputRef.current?.click()}
            loading={loading}
          >
            选择Excel文件
          </Button>
        }
      >
        <input
          type="file"
          ref={fileInputRef}
          onChange={handleFileChange}
          accept=".xlsx, .xls"
          style={{ display: 'none' }}
        />

        {excelData.length > 0 ? (
          <div>
            <Text strong>已选择文件: </Text>
            <Text>{fileName}</Text>
            {/* <div style={{ marginTop: 10 }}>
              <Text>共读取 {excelData.length} 条记录</Text>
              <Table
                dataSource={excelData.slice(0, 5)}
                columns={Object.keys(excelData[0] || {}).map(key => ({
                  title: key,
                  dataIndex: key,
                  width: 150
                }))}
                pagination={false}
                style={{ marginTop: 10 }}
              />
            </div> */}
          </div>
        ) : (
          <div style={{
            border: '2px dashed var(--semi-color-border)',
            borderRadius: 5,
            padding: 40,
            textAlign: 'center'
          }}>
            <Text type="tertiary">请选择Excel文件（.xlsx 或 .xls）</Text>
          </div>
        )}
      </Card>

      {excelData.length > 0 && (
        <>
          <Form
            // @ts-ignore
            labelPosition='top'
            getFormApi={(baseFormApi: BaseFormApi) => formApi.current = baseFormApi}
            onChange={(e) => {
              console.log("设置表格id", e?.values?.table)
              setTableId(e?.values?.table);
            }}
          >
            <Form.Select
              field='table'
              label='选择多维表格sheet表'
              placeholder="请选择sheet表"
              style={{ width: '70%' }}
            >
              {
                Array.isArray(tableMetaList) && tableMetaList.map(({ name, id }) => {
                  return (
                    <Form.Select.Option key={id} value={id}>
                      {name}
                    </Form.Select.Option>
                  );
                })
              }
            </Form.Select>
          </Form>
          <Card
            title="第二步：执行数据对比"
            style={{ marginTop: 20 }}
            headerExtraContent={
              <Button
                theme="solid"
                type="primary"
                onClick={compareData}
                loading={loading}
              >
                开始对比
              </Button>
            }
          >
            <Text>对比Excel数据与当前多维表格数据，识别差异</Text>
          </Card>
        </>

      )}

      {/* {renderFieldMapping()} */}
      {renderComparisonResult}

      {loading && progress > 0 && (
        <div style={{ marginTop: 20 }}>
          <Progress percent={progress} strokeWidth={12} />
          <Text style={{ marginTop: 10, textAlign: 'center' }}>
            正在处理数据，请稍候...
          </Text>
        </div>
      )}
    </div>
  );
};

export default ExcelImportPlugin;