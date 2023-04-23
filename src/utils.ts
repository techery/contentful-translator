import {
  Block,
  BLOCKS,
  Inline,
  INLINES,
  Text,
} from '@contentful/rich-text-types';
import { EntryFields } from 'contentful';
import ExcelJS from 'exceljs';
import path from 'path';

import { ContentModels } from '../../core/schema';
import { Locales } from '../../gen/locales.enum';
import { Localized } from '../models';
import RichText = EntryFields.RichText;

export function getDataFilePath() {
  const dirPath = path.resolve(__dirname, '..', 'data');
  return `${dirPath}/ContentfulData.xlsx`;
}

export function isRichTextValue(payload: any): payload is RichText {
  return payload['nodeType'] === 'document';
}
export function isText(payload: Text | Block | Inline): payload is Text {
  return payload.nodeType === 'text';
}
export function isBlock(payload: Text | Block | Inline): payload is Block {
  return Object.values(BLOCKS).map(String).includes(payload.nodeType);
}
export function isInline(payload: Text | Block | Inline): payload is Inline {
  return Object.values(INLINES).map(String).includes(payload.nodeType);
}

export function parseRichTextField(richText: RichText): string[] {
  function parseBlock(block: Text | Block | Inline): string[] {
    const result: string[] = [];
    if (isText(block)) {
      if (block.value.length === 0) {
        return result;
      }
      return [block.value.trim()];
    }

    if (isBlock(block) || isInline(block)) {
      return block.content.map((x) => parseBlock(x)).flat();
    }

    return [];
  }

  let result: string[] = [];
  for (const block of richText.content) {
    for (const blockData of block.content) {
      result = result.concat(parseBlock(blockData));
    }
  }
  return [...new Set(result)];
}

export function replaceBlock(
  block: Text | Block | Inline,
  baseText: string,
  translatedText: string
) {
  if (isText(block)) {
    if (block.value.length === 0) {
      return;
    }
    if (block.value.trim() === baseText.trim()) {
      block.value = translatedText;
    }
  }

  if (isBlock(block) || isInline(block)) {
    block.content.map((x) => replaceBlock(x, baseText, translatedText)).flat();
  }
}

export function replaceRichTextBlock(
  richText: RichText,
  baseText: string,
  translatedText: string
) {
  const richTextCopy = JSON.parse(JSON.stringify(richText)); // Object.assign({}, richText);
  for (const block of richTextCopy.content) {
    for (const blockData of block.content) {
      replaceBlock(blockData, baseText, translatedText);
    }
  }
  return richTextCopy;
}

export function getValue(
  worksheet: ExcelJS.Worksheet,
  column: number,
  row: number
): string | undefined {
  return worksheet.getRow(row).getCell(column)?.value?.toString();
}

export function getTranslatedValue(
  worksheet: ExcelJS.Worksheet,
  row: number
): Localized<string> {
  const result: Localized<string> = {};
  const headerRow = worksheet.getRow(1);
  for (let index = DataFileColumns.localizations; index <= headerRow.actualCellCount; index++) {
    const locale = Object.values(Locales)
      .map((x) => Object.values(x))
      .flat()
      .find((locale) => headerRow.getCell(index).value === locale);
    if (locale) {
      result[Locales.System.default] = getValue(worksheet, DataFileColumns.defaultText, row) ?? '';
      result[locale] = getValue(worksheet, index, row) ?? '';
    }
  }

  return result;
}

export const DataFileColumns = {
  fieldType: 1,
  defaultText: 2,
  localizations: 3,
}

export const contentTypesBlackList: string[] = [
  ContentModels.SiteRoot,
  ContentModels.ProductPrice,
  ContentModels.ResponsiveImage,
  ContentModels.Warehouse,
  ContentModels.Market,
  ContentModels.CategoryTree,
  ContentModels.ExternalVideo,
  ContentModels.ExternalPage,
  ContentModels.Product,
  ContentModels.ProductVariant,
  'systemMigrationHistory',
  'systemDataMigrationHistory',
];

export function getOrCreateWorksheet(workbook: ExcelJS.Workbook, name: string) {
  const shortenedName = name.slice(0, 30);
  const worksheet = workbook.getWorksheet(shortenedName);
  if (!worksheet) {
    const newWorksheet = workbook.addWorksheet(shortenedName);
    newWorksheet.columns = [
      { header: 'Key', key: 'fieldName' },
      { header: 'English Text', key: 'default' }
    ];

    newWorksheet.getRow(1).font = {
      name: 'Arial Black',
      family: 4,
      bold: true,
    };

    return newWorksheet;
  }
  return worksheet;
}
