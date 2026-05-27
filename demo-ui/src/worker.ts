import { createUniver, LocaleType } from '@univerjs/presets'
import { UniverSheetsCoreWorkerPreset } from '@univerjs/preset-sheets-core/worker'
import { UniverSheetsFilterWorkerPreset } from '@univerjs/preset-sheets-filter/worker'

createUniver({
  locale: LocaleType.ZH_CN,
  locales: {
    zhCN: {},
  },
  presets: [
    UniverSheetsCoreWorkerPreset(),
    UniverSheetsFilterWorkerPreset(),
  ],
})
