import fs from 'node:fs/promises';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';
const root = new URL('../模板/', import.meta.url);
const common = [
  '请按本工作簿填写数据，并将完成后的 .xlsx 文件交回。',
  '保留“数据”工作表名称、第一行表头及列名；数据从第二行开始。不要增加标题行、合并单元格、空表头或把说明写入数据行。',
  '请删除示例数据后填入真实数据；仅使用用户提供的信息，缺失信息请询问，勿编造。',
  '保存为 .xlsx；保留“给AI的提示词”工作表。不要将它设为唯一数据表。',
  '输出前核对姓名和班次编号完全一致，勿修改编号，勿将多个成员写在同一姓名单元格。',
];
const items = [
 ['排班名单模板', [['姓名','电话','宿舍','班级','专业','辅导员'],['张三','13800000001','1-101','一班','示例专业','李老师'],['李四','13800000002','1-102','','',''],['王五','13800000003','2-103','','',''],['赵六','13800000004','2-104','','','']], '每人一行；姓名必填且唯一。电话、宿舍、班级、专业、辅导员可留空。电话号码按文本保存。'],
 ['班次设置模板', [['班次编号','日期','班次名称','开始时间','结束时间','人数'],['1','2026-11-02','中午','12:00','13:00',2],['2','2026-11-02','下午','14:00','16:00',2],['3','2026-11-02','晚上','18:00','20:30',2],['4','2026-11-03','中午','12:00','13:00',2]], '每个实际班次一行，编号唯一；日期使用 YYYY-MM-DD，时间使用 HH:MM，人数为正整数。结束时间不晚于开始时间表示次日结束。'],
 ['时间冲突模板', [['姓名','冲突班次编号'],['张三','1,3'],['李四','2']], '每行填写一人及其不能参加的班次编号，多个编号用逗号分隔。姓名和编号必须与名单、班次设置表一致。无冲突时仅保留表头。']
];
for (const [name, rows, extra] of items) {
 const wb=Workbook.create(); const data=wb.worksheets.add('数据'); const notes=wb.worksheets.add('给AI的提示词');
 data.getRangeByIndexes(0,0,rows.length,rows[0].length).values=rows;
 data.getRangeByIndexes(0,0,1,rows[0].length).format.font.bold=true;
 data.getRangeByIndexes(0,0,rows.length,rows[0].length).format.columnWidth=18;
 data.freezePanes.freezeRows(1);
 notes.getRange('A1').values=[['给 AI 的填写提示词（复制下列内容交给 AI）']];
 notes.getRange('A1').format.font.bold=true;
 notes.getRange('A2:A7').values=[...common,extra].map(x=>[x]);
 notes.getRange('A1:A7').format.columnWidth=100;
 notes.getRange('A1:A7').format.wrapText=true;
 notes.getRange('A1:A7').format.autofitRows();
 const file=await SpreadsheetFile.exportXlsx(wb); await file.save(decodeURIComponent(new URL(name+'.xlsx',root).pathname).replace(/^\/(\w:)/,'$1'));
}
