import { BACKUP_FORMAT, BACKUP_VERSION } from "../config.js";

export function createBackup(records, { exportedAt, bankVersion }) {
  return `${JSON.stringify({ format: BACKUP_FORMAT, version: BACKUP_VERSION, bankVersion, exportedAt, records }, null, 2)}\n`;
}

export function parseBackup(text, currentBankVersion) {
  let value;
  try { value = JSON.parse(text); } catch { throw new Error("备份文件不是有效 JSON"); }
  if (value.format !== BACKUP_FORMAT || value.version !== BACKUP_VERSION) throw new Error("备份格式或版本不受支持");
  if (!Number.isInteger(value.bankVersion) || value.bankVersion > currentBankVersion) throw new Error("备份题库版本高于当前应用");
  if (!Array.isArray(value.records) || value.records.some(record => !record || typeof record.questionId !== "string")) {
    throw new Error("备份记录结构无效");
  }
  return value;
}
