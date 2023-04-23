import { EntryFields } from 'contentful';
import * as cfManagement from 'contentful-management';
import { Environment } from 'contentful-management';
import ExcelJS from 'exceljs';

import config from '../../config';
import { Locales } from '../../gen';
import { Localized } from '../models';
import {
  contentTypesBlackList,
  getDataFilePath,
  getOrCreateWorksheet,
  getTranslatedValue,
  isRichTextValue,
  parseRichTextField,
  replaceRichTextBlock,
} from './utils';
import RichText = EntryFields.RichText;

async function exportAllLabels() {
  console.log(`Export started from ${config.environment}`);

  const workbook = new ExcelJS.Workbook();
  const knownKeys: string[] = [];
  const loader = new ContentfulItemsLoader();

  async function exportValue(entry: cfManagement.Entry, field: string) {
    const defaultFieldValue = entry.fields[field][Locales.System.default];

    if (isRichTextValue(defaultFieldValue)) {
      if (defaultFieldValue.content.length === 0) {
        return;
      }
      for (const value of parseRichTextField(defaultFieldValue)) {
        if (!knownKeys.includes(value)) {
          getOrCreateWorksheet(workbook, entry.sys.contentType.sys.id).addRow({
            default: value,
            fieldName: field
          });

          knownKeys.push(value);
        }
      }
      return;
    }

    if (typeof defaultFieldValue !== 'string') {
      return;
    }

    if (knownKeys.find((x) => x === defaultFieldValue.trim())) {
      return;
    }

    getOrCreateWorksheet(workbook, entry.sys.contentType.sys.id).addRow({
      default: defaultFieldValue.trim(),
      fieldName: field
    });

    knownKeys.push(defaultFieldValue.trim());
  }

  await loader.loadLocalizableEntries(exportValue);

  const path = getDataFilePath()
  await workbook.xlsx.writeFile(path);

  console.log(`The data was successfully exported to ${path}`);
}

async function importAllLabels() {
  console.log(`Import started to ${config.environment}`);

  const filePath = getDataFilePath();
  const workbook = new ExcelJS.Workbook();
  const content = await workbook.xlsx.readFile(filePath);

  const localizations: Localized<string>[] = [];
  for (const sheet of content.worksheets) {
    for (let row = 2; row < sheet.rowCount; row++) {
      localizations.push(getTranslatedValue(sheet, row));
    }
  }

  console.log(`Processed the data file`);

  async function importValue(entry: cfManagement.Entry, field: string) {
    const defaultFieldValue = entry.fields[field][Locales.System.default];

    if (isRichTextValue(defaultFieldValue)) {
      if (defaultFieldValue.content.length === 0) {
        return;
      }
      for (const value of parseRichTextField(defaultFieldValue)) {
        const localization = localizations.find(
          (x) => x[Locales.System.default] === value
        );

        if (localization) {
          const localizedObjects = Object.keys(localization).map((locale) => {
            const localizedRichText = (entry.fields[field][locale] ||
              defaultFieldValue) as RichText;
            return {
              [locale]: replaceRichTextBlock(
                localizedRichText,
                value,
                localization[locale]
              ),
            };
          });
          const localizedField = localizedObjects.reduce((obj, item) => {
            return {
              ...obj,
              ...item,
            };
          });
          localizedField[Locales.System.default] =
            entry.fields[field][Locales.System.default];

          entry.fields[field] = {
            ...localizedField,
          };
        }
      }
      return;
    }
    const localization = localizations.find(
      (x) => `${x[Locales.System.default]}`.trim() === `${defaultFieldValue}`.trim()
    );
    if (localization) {
      entry.fields[field] = {
        ...localization,
        [Locales.System.default]: entry.fields[field][Locales.System.default],
      };
    }
  }

  async function saveEntry(entry: cfManagement.Entry) {
    try {
      await entry.update().then((x) => x.publish());
    } catch (e) {
      console.log(
        `Failed to update item ${entry.sys.id} : ${entry.sys.contentType.sys.id} : ${e}`
      );
    }
  }

  const loader = new ContentfulItemsLoader();
  await loader.loadLocalizableEntries(importValue, saveEntry);

  console.log(`The data was successfully imported`);
}

class ContentfulItemsLoader {
  client = cfManagement.createClient({
    accessToken: config.accessToken,
  });

  private async loadEntries(
    env: Environment,
    offset: number,
    offsetStep: number
  ) {
    console.log(`Fetching entries from ${offset} to ${offset + offsetStep}`);
    return env.getEntries({
      skip: offset,
      limit: offsetStep,
    });
  }

  async validateEntry(
    env: cfManagement.Environment,
    item: cfManagement.Entry,
    field: string
  ) {
    if (contentTypesBlackList.includes(item.sys.contentType.sys.id)) {
      return false;
    }
    const contentType = await env.getContentType(item.sys.contentType.sys.id);

    const fieldModel = contentType.fields.find(
      (fieldModel) => fieldModel.id === field
    );
    if (!fieldModel) {
      throw new Error(
        `Field with name ${fieldModel} is not found in model ${contentType.name}`
      );
    }
    if (!fieldModel.localized) {
      return false;
    }
    const defaultValue = item.fields[field][Locales.System.default];

    if (!defaultValue || defaultValue['sys']) {
      return false;
    }

    if (Array.isArray(defaultValue)) {
      return;
    }

    if (typeof defaultValue === 'string') {
      return defaultValue.trim().length !== 0;
    }

    return !!item.fields[field][Locales.System.default];
  }

  async loadLocalizableEntries(
    fieldProcessCallback: (
      entry: cfManagement.Entry,
      field: string
    ) => Promise<void>,
    finalizeCallback?: (entry: cfManagement.Entry) => Promise<void>
  ) {
    const env = await this.client
      .getSpace(config.spaceId)
      .then((space) => space.getEnvironment(config.environment));

    let offset = 0;
    const offsetStep = 200;

    let entities = await this.loadEntries(env, offset, offsetStep);

    do {
      offset += offsetStep;

      for (const item of entities.items) {
        let shouldFinalize = false;
        for (const field of Object.keys(item.fields)) {
          if (!(await this.validateEntry(env, item, field))) {
            continue;
          }
          shouldFinalize = true;
          await fieldProcessCallback(item, field);
        }
        if (finalizeCallback && shouldFinalize) {
          await finalizeCallback(item);
        }
      }
      entities = await this.loadEntries(env, offset, offsetStep);
    } while (entities.items.length > 0);
  }
}

if (process.argv.length < 3) {
  throw new Error('Specify the operation type. Valid types: import | export');
}

const operationType = process.argv[2];

if (operationType === 'import') {
  importAllLabels();
} else if (operationType === 'export') {
  exportAllLabels();
} else {
  throw new Error(
    `Invalid the operation type. Valid types: import | export. Received ${process.argv[2]}`
  );
}
