import {
  BooleanNumber,
  createUniver,
  defaultTheme,
  LocaleType,
  LogLevel,
  mergeLocales,
  SheetTypes,
  UniverInstanceType,
} from '@univerjs/presets'

import { CalculationMode, UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import sheetsCoreEnUs from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'

import { UniverSheetsThreadCommentPreset } from '@univerjs/preset-sheets-thread-comment'
import sheetsThreadCommentEnUs from '@univerjs/preset-sheets-thread-comment/locales/en-US'
import '@univerjs/preset-sheets-thread-comment/lib/index.css'

import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import sheetsConditionalFormattingEnUs from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'

import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import sheetsDataValidationEnUs from '@univerjs/preset-sheets-data-validation/locales/en-US'
import '@univerjs/preset-sheets-data-validation/lib/index.css'

import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import sheetsDrawingEnUs from '@univerjs/preset-sheets-drawing/locales/en-US'
import '@univerjs/preset-sheets-drawing/lib/index.css'

import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import sheetsFilterEnUs from '@univerjs/preset-sheets-filter/locales/en-US'
import '@univerjs/preset-sheets-filter/lib/index.css'

import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import sheetsFindReplaceEnUs from '@univerjs/preset-sheets-find-replace/locales/en-US'
import '@univerjs/preset-sheets-find-replace/lib/index.css'

import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link'
import sheetsHyperLinkEnUs from '@univerjs/preset-sheets-hyper-link/locales/en-US'
import '@univerjs/preset-sheets-hyper-link/lib/index.css'

import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import sheetsSortEnUs from '@univerjs/preset-sheets-sort/locales/en-US'
import '@univerjs/preset-sheets-sort/lib/index.css'

import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import sheetsNoteEnUs from '@univerjs/preset-sheets-note/locales/en-US'
import '@univerjs/preset-sheets-note/lib/index.css'

import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import sheetsTableEnUs from '@univerjs/preset-sheets-table/locales/en-US'
import '@univerjs/preset-sheets-table/lib/index.css'

import { UniverSheetsZenEditorPlugin } from '@univerjs/sheets-zen-editor'
import sheetsZenEditorEnUs from '@univerjs/sheets-zen-editor/locale/en-US'
import '@univerjs/sheets-zen-editor/lib/index.css'

import { UniverSheetsCrosshairHighlightPlugin } from '@univerjs/sheets-crosshair-highlight'
import sheetsCrosshairHighlightEnUs from '@univerjs/sheets-crosshair-highlight/locale/en-US'
import '@univerjs/sheets-crosshair-highlight/lib/index.css'

// oxlint-disable-next-line import/default
import workerURL from './worker.ts?worker&url'

// import { setupUniverDebugPlugin } from './plugins/debug'

export function setupUniver() {
  const { univerAPI, univer } = createUniver({
    locale: LocaleType.EN_US,
    locales: {
      [LocaleType.EN_US]: mergeLocales(
        sheetsCoreEnUs,
        sheetsThreadCommentEnUs,
        sheetsConditionalFormattingEnUs,
        sheetsDataValidationEnUs,
        sheetsDrawingEnUs,
        sheetsFilterEnUs,
        sheetsFindReplaceEnUs,
        sheetsHyperLinkEnUs,
        sheetsSortEnUs,
        sheetsNoteEnUs,
        sheetsTableEnUs,
        sheetsZenEditorEnUs,
        sheetsCrosshairHighlightEnUs,
      ),
    },
    logLevel: LogLevel.VERBOSE,
    theme: defaultTheme,
    presets: [
      UniverSheetsCorePreset({
        container: 'univer',
        header: true,
        toolbar: false,
        workerURL: new Worker(new URL(workerURL, import.meta.url), {
          type: 'module',
        }),
        formula: {
          initialFormulaComputing: CalculationMode.FORCED,
        },
        // footer: {
        //   addSheetButtonConfig: {
        //     // show: false,
        //     defaultRowCount: 10,
        //     defaultColumnCount: 5,
        //   },
        // },
      }),
      UniverSheetsDrawingPreset({
        collaboration: false,
        // allowImageSize: 0.01 * 1024 * 1024, // 10KB
      }),
      UniverSheetsThreadCommentPreset({
        collaboration: false,
      }),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsDataValidationPreset(),
      UniverSheetsFilterPreset({
        enableSyncSwitch: true,
      }),
      UniverSheetsFindReplacePreset(),
      UniverSheetsSortPreset(),
      UniverSheetsNotePreset(),
      UniverSheetsTablePreset(),
      UniverSheetsHyperLinkPreset(),
    ],
    plugins: [UniverSheetsCrosshairHighlightPlugin, UniverSheetsZenEditorPlugin],
  })

  // setupUniverDebugPlugin(univer)

  univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: 'workbook-01',
    locale: LocaleType.EN_US,
    name: 'Untitled spreadsheet',
    appVersion: '0.24.0',
    sheetOrder: ['sheet-01'],
    sheets: {
      'sheet-01': {
        type: SheetTypes.GRID,
        id: 'sheet-01',
        name: 'Sheet1',
        cellData: {},
        hidden: BooleanNumber.FALSE,
        rowCount: 1000,
        columnCount: 20,
        zoomRatio: 1,
        scrollTop: 0,
        scrollLeft: 0,
        defaultColumnWidth: 93,
        defaultRowHeight: 27,
        status: 1,
        showGridlines: 1,
        hideRow: [],
        hideColumn: [],
        rowHeader: {
          width: 46,
          hidden: BooleanNumber.FALSE,
        },
        columnHeader: {
          height: 20,
          hidden: BooleanNumber.FALSE,
        },
        selections: ['A1'],
        rightToLeft: BooleanNumber.FALSE,
        pluginMeta: {},
      },
    },
  })

  return univerAPI
}
